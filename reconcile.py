"""
FMX FCA reconciliation tool.

Combines a technician capture file and (optionally) an engineering partner's
equipment list into one list in FMX's Capital Planner import format.
Values with a clear right answer are fixed automatically; anything the tool
can't resolve with confidence is marked needs_review with a plain-language reason.

Run:  python reconcile.py            (uses the files in data/)
      python reconcile.py --tech path/to/other_technician_file.xlsx
See README.md for details.
"""

import argparse
import math
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
RULES = HERE / "rules"
THIS_YEAR = date.today().year
EARLIEST_YEAR = 1900

# Possible column headers for each field (compared lowercase, letters/numbers only).
HEADER_ALIASES = {
    "building": ["bldg", "building", "facility", "site"],
    "asset_raw": ["equiptype", "equipmenttype", "equipmentdescription", "assettype", "equipment", "description"],
    "manufacturer": ["mfr", "manufacturer", "make", "mfg"],
    "model": ["model", "modelnumber", "modelno"],
    "serial": ["serial", "serialnumber", "serialno", "sn"],
    "year_raw": ["yrinstalled", "yearinstalled", "installyear", "installed", "yearofmanufacture", "year"],
    "condition_raw": ["cond110", "cond", "condition", "conditionrating"],
    "location": ["location", "roomarea", "room", "area"],
    "notes": ["notes", "note", "comments", "comment"],
    "photo": ["photo", "photos", "image"],
    "tag": ["tag", "assettag"],
}

PARTNER_CONDITION = {"good": 4, "fair": 3, "poor": 2, "critical": 1}

TARGET_COLUMNS = ["building", "asset_type", "manufacturer", "model", "serial_number",
                  "install_year", "condition", "location", "source", "status", "review_reason"]


# ---------------------------------------------------------------- helpers
def key(text):
    """Uppercase letters and numbers only, for matching spellings."""
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def words(text):
    """Lowercase words with punctuation and (parentheses) removed."""
    text = re.sub(r"\(.*?\)", " ", str(text).lower())
    return " ".join(re.sub(r"[^a-z0-9]+", " ", text).split())


def clean(value):
    """Trimmed string, or '' for blanks."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def load_alias_table(filename, value_col, keyfunc):
    table = pd.read_csv(RULES / filename, dtype=str)
    return {keyfunc(a): v.strip() for a, v in zip(table["alias"], table[value_col])}


def read_any(path, **kwargs):
    path = Path(path)
    if not path.exists():
        sys.exit(f"Can't find file: {path}")
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=str, **kwargs)
    return pd.read_csv(path, dtype=str, **kwargs)


# ---------------------------------------------------------------- loading
def load_source(path, source):
    """Read a file as provided, find its header row, map columns.
    Returns (data rows, skipped non-data rows)."""
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        # Use the sheet whose header looks most like equipment data (skips Legend tabs).
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
    else:
        sheets = {"csv": pd.read_csv(path, header=None, dtype=str, keep_default_na=False, na_values=[""])}

    best = None
    for name, raw in sheets.items():
        for i in range(min(15, len(raw))):
            mapping = map_headers(raw.iloc[i].tolist())
            if best is None or len(mapping) > len(best[2]):
                best = (name, i, mapping, raw)
    name, header_idx, mapping, raw = best
    if "asset_raw" not in mapping.values() or "building" not in mapping.values():
        sys.exit(f"Couldn't find the equipment columns in {path.name}. Found headers: {raw.iloc[header_idx].tolist()}")

    header_values = [key(v) for v in raw.iloc[header_idx].tolist()]
    rows, skipped = [], []
    for i in range(header_idx + 1, len(raw)):
        values = raw.iloc[i].tolist()
        excel_row = i + 1
        filled = [clean(v) for v in values if clean(v)]
        if not filled:
            continue  # completely blank spacer row
        if [key(v) for v in values] == header_values:
            skipped.append((source, excel_row, "Repeated header row"))
            continue
        rec = {field: clean(values[col]) for col, field in mapping.items()}
        if len(filled) == 1 and not rec.get("asset_raw"):
            skipped.append((source, excel_row, f"Note/footer text, not equipment: \"{filled[0][:60]}\""))
            continue
        rec.update(source=source, rows=[f"{source} row {excel_row}"])
        rows.append(rec)
    return rows, skipped


def map_headers(header_cells):
    mapping = {}
    for col, cell in enumerate(header_cells):
        k = key(cell).lower()
        for field, aliases in HEADER_ALIASES.items():
            if k in aliases and field not in mapping.values():
                mapping[col] = field
                break
    return mapping


# ---------------------------------------------------------------- normalizing
class Normalizer:
    def __init__(self, asset_types_path, target_path):
        types = read_any(asset_types_path)
        self.asset_types = [t.strip() for t in types["asset_type"].dropna()]
        self.type_by_words = {words(t): t for t in self.asset_types}
        self.type_synonyms = {words(a): t for a, t in
                              load_alias_table("asset_type_synonyms.csv", "asset_type", words).items()}
        bad = set(self.type_synonyms.values()) - set(self.asset_types)
        if bad:
            sys.exit(f"rules/asset_type_synonyms.csv points to types FMX doesn't accept: {bad}")

        target = read_any(target_path)
        self.buildings = sorted(set(target["building"].dropna().str.strip()))
        self.building_aliases = load_alias_table("building_aliases.csv", "building", key)
        self.mfr_aliases = load_alias_table("manufacturer_aliases.csv", "manufacturer", key)
        kw = pd.read_csv(RULES / "review_keywords.csv", dtype=str)
        self.keywords = list(zip(kw["keyword"].str.lower(), kw["reason"]))

    def building(self, raw):
        if not raw:
            return "", "Building is blank"
        b = self.building_aliases.get(key(raw))
        if b in self.buildings:
            return b, None
        return "", f"Building \"{raw}\" doesn't match a known building"

    def asset_type(self, raw):
        if not raw:
            return "", "Asset type is blank"
        if re.search(r"\bqty\b|\(\s*\d+\s*\)|typical of", raw, re.I):
            return "", f"Row \"{raw}\" covers several assets; split into one row per asset"
        # remove a trailing tag (AHU-1) and anything after a comma (sizes like 10 ton)
        text = words(re.sub(r",.*$", "", raw))
        text = re.sub(r"\b[a-z]{1,4} \d+[a-z]?\b$", "", text).strip()
        if text in self.type_by_words:
            return self.type_by_words[text], None
        if text in self.type_synonyms:
            return self.type_synonyms[text], None
        # official name contained in a longer description, e.g. "Base-Mounted Pump" -> Pump
        hits = [t for w, t in self.type_by_words.items() if re.search(rf"\b{re.escape(w)}\b", text)]
        hits += [t for w, t in self.type_synonyms.items() if len(w) > 4 and re.search(rf"\b{re.escape(w)}\b", text)]
        hits = sorted(set(hits), key=len, reverse=True)
        if len(hits) == 1:
            return hits[0], None
        return "", f"Asset type \"{raw}\" isn't on FMX's list or is too vague to map"

    def manufacturer(self, raw):
        if not raw:
            return "", None
        m = self.mfr_aliases.get(key(raw))
        if m:
            return m, None
        return raw, f"Manufacturer \"{raw}\" isn't in the known spellings list; confirm and add it to rules/manufacturer_aliases.csv"

    @staticmethod
    def year(raw):
        """Returns (year or '', reason or None)."""
        v = raw.strip().lower()
        if v in ("", "unknown", "unk", "n/a", "na", "?"):
            return "", None
        two = lambda yy: 2000 + yy if yy <= THIS_YEAR % 100 else 1900 + yy
        # approximate single years are accepted (see README)
        v = re.sub(r"^(approximately|approx\.?|circa|about|est\.?|ca\.?|c\.?|~)\s*", "", v)
        v = re.sub(r"^mfg\s*(\d{1,2}/)?", "", v)          # manufacture date, e.g. Mfg 04/2011
        v = v.rstrip("?").strip()
        if re.fullmatch(r"\d{4}", v):
            y = int(v)
        elif re.fullmatch(r"'?\d{2}", v):
            y = two(int(v.lstrip("'")))
        else:
            return "", f"Install year \"{raw}\" isn't a single year"
        if not EARLIEST_YEAR <= y <= THIS_YEAR:
            return "", f"Install year \"{raw}\" isn't a realistic year"
        return str(y), None

    @staticmethod
    def condition(raw, source):
        """Returns (1-5 or '', reason or None)."""
        v = raw.strip().lower()
        if not v:
            return "", "Condition is blank"
        if source == "partner" or v in PARTNER_CONDITION:
            if v in PARTNER_CONDITION:
                if source == "technician":
                    return "", f"Condition \"{raw}\" uses words instead of the 1 to 10 scale"
                return str(PARTNER_CONDITION[v]), None
            return "", f"Condition \"{raw}\" isn't Good, Fair, Poor, or Critical"
        try:
            n = float(v)
        except ValueError:
            return "", f"Condition \"{raw}\" isn't a number from 1 to 10"
        if not 1 <= n <= 10:
            return "", f"Condition \"{raw}\" is outside the 1 to 10 scale"
        return str(math.ceil(n / 2)), None

    def note_flags(self, text):
        t = text.lower()
        return [reason for kw, reason in self.keywords if kw in t]


def extract_tag(rec):
    """Equipment tag like RTU-4, P-1, LP-2A from the tag column or the notes."""
    if rec.get("tag"):
        m = re.match(r"\s*([A-Za-z]{1,4}-\d+[A-Za-z]?)\b", rec["tag"])
        if m and "to" not in rec["tag"].lower():
            return m.group(1).upper()
    m = re.search(r"\b([A-Za-z]{1,4}-\d+[A-Za-z]?)\b", rec.get("notes", ""))
    return m.group(1).upper() if m else ""


def normalize(rec, n):
    rec["reasons"] = []
    add = lambda r: r and rec["reasons"].append(r)
    rec["building"], r = n.building(rec.get("building", "")); add(r)
    rec["asset_type"], r = n.asset_type(rec.get("asset_raw", "")); add(r)
    rec["manufacturer"], r = n.manufacturer(rec.get("manufacturer", "")); add(r)
    rec["model"] = rec.get("model", "").upper()
    rec["serial_number"] = rec.get("serial", "")
    rec["install_year"], r = n.year(rec.get("year_raw", "")); add(r)
    rec["approx_year"] = bool(rec["install_year"]) and bool(re.search(r"approx|~|circa|\bc\.|\?|about|est", rec.get("year_raw", ""), re.I))
    rec["condition"], r = n.condition(rec.get("condition_raw", ""), rec["source"]); add(r)
    rec["location"] = rec.get("location", "")
    rec["tag"] = extract_tag(rec)
    rec["serial_key"] = key(rec["serial_number"])
    if rec["serial_key"] in ("SEESCHEDULE", "NA", "NONE", "UNKNOWN"):
        rec["serial_key"] = ""
        add("Serial number is not on the nameplate list (\"see schedule\")")
               # Excel turns long numbers into scientific notation (4.43E+09) and drops the real digits
    if re.fullmatch(r"\d{6,}0000", rec["serial_key"]):
        rec["serial_key"] = ""
        add(f"Serial number \"{rec['serial_number']}\" looks rounded by Excel; check the nameplate photo")
    for r in n.note_flags(rec.get("notes", "")):
        add(r)
    return rec


def nameplate_check(rec):
    if not rec["manufacturer"] and not rec["model"] and not rec["serial_number"]:
        photo = f"; technician photo {rec['photo']} on file" if rec.get("photo") else ""
        rec["reasons"].append(f"Nameplate data missing{photo}")


# ---------------------------------------------------------------- duplicates
def same_asset(a, b):
    if a["building"] and a["building"] != b["building"]:
        return False
    if a["serial_key"] and b["serial_key"]:
        return a["serial_key"] == b["serial_key"]
    return bool(a["tag"]) and a["tag"] == b["tag"]


def merge(keep, other, label_keep, label_other):
    """Merge `other` into `keep`. Blanks are filled; real disagreements are flagged."""
    for f in ("asset_type", "manufacturer", "model", "serial_number", "location"):
        if not keep[f]:
            keep[f] = other[f]
    if keep["asset_type"] and other["asset_type"] and keep["asset_type"] != other["asset_type"]:
        keep["reasons"].append(f"{label_keep} says {keep['asset_type']}, {label_other} says {other['asset_type']}")
    if keep["model"] and other["model"] and key(keep["model"]) != key(other["model"]):
        keep["reasons"].append(f"Model differs: {label_keep} {keep['model']}, {label_other} {other['model']}")

    # install year: keep's year wins; small gaps are expected (manufacture vs. install)
    ya, yb = keep["install_year"], other["install_year"]
    if not ya:
        keep["install_year"] = yb
    elif yb and abs(int(ya) - int(yb)) > 2:
        keep["reasons"].append(f"Install year differs: {label_keep} {ya}, {label_other} {yb}")

    # condition: within 1 point keep the worse; 2+ points apart goes to review
    ca, cb = keep["condition"], other["condition"]
    if not ca:
        keep["condition"] = cb
    elif cb:
        if abs(int(ca) - int(cb)) >= 2:
            keep["reasons"].append(
                f"Condition differs: {label_keep} rated {keep['condition_raw']} ({ca} of 5), "
                f"{label_other} rated {other['condition_raw']} ({cb} of 5)")
        keep["condition"] = str(min(int(ca), int(cb)))

    # carry over other row's flags, except ones the merge just resolved
    resolved = ("Nameplate data missing", "Condition is blank", "Asset type is blank")
    for r in other["reasons"]:
        if r not in keep["reasons"]:
            keep["reasons"].append(r)
    keep["reasons"] = [r for r in keep["reasons"]
                       if not (r.startswith(resolved) and keep["serial_number"] and keep["condition"])]
    keep["rows"] += other["rows"]
    keep["notes"] = "; ".join(x for x in (keep.get("notes", ""), other.get("notes", "")) if x)


def dedupe_within(rows, label):
    out, merged = [], 0
    for rec in rows:
        match = next((k for k in out if same_asset(k, rec)), None)
        if match:
            merge(match, rec, f"{label} first entry", f"{label} repeat entry")
            merged += 1
        else:
            out.append(rec)
    return out, merged


def match_sources(tech, partner):
    merged = 0
    used = set()
    for p in partner:
        cands = [t for t in tech if id(t) not in used and same_asset(t, p)]
        # tag matched but both have different serials: not the same unit
        cands = [t for t in cands if not (t["serial_key"] and p["serial_key"] and t["serial_key"] != p["serial_key"])]
        if not cands:
            p["source"] = "partner"
            continue
        t = cands[0]
        used.add(id(t))
        # nameplate fields: partner surveyed nameplates, so prefer theirs; location too (cleaner)
        for f in ("manufacturer", "model", "serial_number", "location", "tag"):
            if p[f]:
                t[f] = p[f]
        merge(t, p, "technician", "partner")
        t["source"] = "both"
        p["merged_into"] = t
        merged += 1
    return [p for p in partner if "merged_into" not in p], merged


# ---------------------------------------------------------------- main
def run(args):
    n = Normalizer(args.asset_types, args.target)
    tech, skipped = load_source(args.tech, "technician")
    partner, skipped_p = ([], []) if not args.partner else load_source(args.partner, "partner")
    skipped += skipped_p
    rows_in = {"technician": len(tech), "partner": len(partner)}

    for rec in tech + partner:
        normalize(rec, n)
    tech, tech_dupes = dedupe_within(tech, "technician")
    partner, partner_dupes = dedupe_within(partner, "partner")
    partner_left, cross = match_sources(tech, partner)
    final = tech + partner_left
    for rec in final:
        nameplate_check(rec)
        # keep the equipment tag (LP-1A, RTU-4) so identical-looking units can be told apart
        if rec["tag"] and rec["tag"] not in rec["location"].upper():
            rec["location"] = f"{rec['location']} ({rec['tag']})".strip()
        rec["reasons"] = list(dict.fromkeys(rec["reasons"]))
        rec["status"] = "needs_review" if rec["reasons"] else "ready"
        rec["review_reason"] = "; ".join(rec["reasons"])

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    order = {b: i for i, b in enumerate(n.buildings)}
    final.sort(key=lambda r: (order.get(r["building"], 99), r["asset_type"], r["location"]))
    df = pd.DataFrame([{c: r.get(c, "") for c in TARGET_COLUMNS} for r in final])
    df.to_csv(out / "import_ready.csv", index=False)

    review = pd.DataFrame([{**{c: r.get(c, "") for c in TARGET_COLUMNS},
                            "original_asset_type": r.get("asset_raw", ""),
                            "original_install_year": r.get("year_raw", ""),
                            "original_condition": r.get("condition_raw", ""),
                            "notes": r.get("notes", ""), "photo": r.get("photo", ""),
                            "source_rows": "; ".join(r["rows"])}
                           for r in final if r["status"] == "needs_review"])
    review.to_csv(out / "review_list.csv", index=False)

    write_summary(out, args, rows_in, final, tech_dupes + partner_dupes, cross, skipped)
    print((out / "run_summary.txt").read_text())


def write_summary(out, args, rows_in, final, within, cross, skipped):
    from collections import Counter
    total_in = sum(rows_in.values())
    ready = sum(r["status"] == "ready" for r in final)
    flagged = len(final) - ready
    consumed = sum(len(r["rows"]) for r in final)
    approx = sum(r.get("approx_year", False) for r in final)
    reasons = Counter()
    for r in final:
        for reason in r["reasons"]:
            label = re.sub(r";.*|\(.*?\)", "", reason)       # drop photo names and examples
            label = " ".join(re.sub(r'"[^"]*"', "", label).split(":")[0].split())
            reasons[label] += 1

    lines = [
        "FMX reconciliation run summary",
        f"Technician file: {args.tech}",
        f"Partner file:    {args.partner or '(none)'}",
        "",
        f"Rows in:            {total_in}  (technician {rows_in['technician']}, partner {rows_in['partner']})",
        f"Rows out:           {len(final)}",
        f"  Ready:            {ready}",
        f"  Needs review:     {flagged}",
        f"Duplicates merged:  {within + cross}  (technician vs partner {cross}, repeat entries in the same file {within})",
        "",
        f"Approximate install years kept: {approx}  (e.g. \"approx. 2004\" became 2004; the capital plan only needs a rough age)",
        "",
        f"Check: {total_in} rows in = {len(final)} rows out + {within + cross} merged "
        f"-> {'OK, no rows lost' if total_in == len(final) + within + cross == consumed else 'MISMATCH, investigate'}",
        "",
        "Review reasons (a row can have more than one):",
    ]
    lines += [f"  {c:>3}  {r}" for r, c in reasons.most_common()]
    if skipped:
        lines += ["", "Non-equipment rows skipped (not counted as rows in):"]
        lines += [f"  {s} row {row}: {why}" for s, row, why in skipped]
    (out / "run_summary.txt").write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Reconcile FCA technician and partner files into FMX import format.")
    ap.add_argument("--tech", default=HERE / "data" / "Technician_Capture.csv.xlsx", help="technician capture file (.xlsx or .csv)")
    ap.add_argument("--partner", default=HERE / "data" / "Partner_Equipment_List.xlsx", help="partner equipment list (.xlsx or .csv); use --partner '' to skip")
    ap.add_argument("--asset-types", default=HERE / "data" / "Asset_Types.xlsx")
    ap.add_argument("--target", default=HERE / "data" / "Target_Format.xlsx")
    ap.add_argument("--out", default=HERE / "output")
    run(ap.parse_args())
