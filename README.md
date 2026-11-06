# FMX FCA Reconciliation Tool

Combines a technician's capture file and an engineering partner's equipment list into one clean list in FMX's Capital Planner import format. It fixes the values that have a clear right answer and sends everything else to a person with a plain-language reason. No row is lost.

## Setup (one time, about 2 minutes)
You need Python 3.10 or newer.

```
cd fmx-reconcile
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```
(`streamlit` is only needed for the web page.)

## Run it
Put the files in `data/`, then:

```
python reconcile.py
```

Use different files (any name, `.xlsx` or `.csv`):

```
python reconcile.py --tech path/to/new_technician_file.xlsx --partner path/to/partner_list.xlsx
python reconcile.py --tech path/to/new_technician_file.csv --partner ""     # no partner file
```

Options: `--asset-types` (FMX's accepted types), `--target` (import format, used for the exact building names), `--out` (output folder, default `output/`).

## Or use the web page
```
streamlit run app.py
```
A page opens in your browser: upload the technician file (and the partner file if there is one), click **Run**, review the flagged rows, and download the three outputs. It uses the same logic as `reconcile.py`.

## What you get (in `output/`)
| File | What it is |
|---|---|
| `import_ready.csv` | Every asset in FMX's import format, with `status` = ready or needs_review |
| `review_list.csv` | Only the needs_review rows, with the reason, the original values, notes, photo, and the source row numbers so you can find them in the original files |
| `run_summary.txt` | Rows in, ready, flagged, duplicates merged, a count of each review reason, and a check that no row was lost |

## Teaching it a new spelling (no code)
Edit the files in `rules/` and run again:
- `asset_type_synonyms.csv`: free text to an FMX asset type (e.g., `RTU,Rooftop Unit`)
- `manufacturer_aliases.csv`: spellings to one manufacturer name. Case, spaces, and punctuation are ignored, so `Square-D`, `SquareD`, and `square d` all match `Square D`
- `building_aliases.csv`: building nicknames to the exact building name
- `review_keywords.csv`: words in notes that should send a row to review

## Decisions the tool makes
**Reading files as provided.** Inputs arrived as `.xlsx` (the instructions said CSV); the tool reads either. It finds the header row on its own (the partner file has title rows above it), skips blank rows, and lists repeated header rows and footer notes in the summary instead of silently dropping them.

**Condition to 1 to 5.**
- Technician 1 to 10: divide by 2 and round up (1-2 = 1, 3-4 = 2, 5-6 = 3, 7-8 = 4, 9-10 = 5; 3.5 = 2).
- Partner: Good = 4, Fair = 3, Poor = 2, Critical = 1. Good is 4, not 5, because FMX's 5 means new or like new, and the partner's legend defines Good as "routine maintenance only," not new. Critical ("failed or at risk of failure") = 1.
- Blank, out of range (like 11), or the wrong scale (a word in the 1 to 10 column) goes to review.

**Install year.**
- `'04`, `04` = 2004; `'98` = 1998 (two digits up to this year are 2000s, otherwise 1900s).
- Approximate single years (`approx. 2004`, `~2004`, `c. 2012`, `2004?`) are kept as that year and marked ready. Capital planning uses the year to estimate remaining life, and a year or two either way doesn't change that decision; flagging all of them would bury the reviewer.
To keep the guesses visible, the run summary counts how many install years were approximate.
- `Mfg 04/2011` uses 2011. The partner's file also gives year of manufacture, not install, so this is treated the same way.
- `unknown` or blank = blank (allowed by the format).
- Anything that isn't one year (`2000s`), is a typo (`19987`), or is in the future (`2031`) goes to review.

**Duplicates.**
- Same building plus the same serial number (ignoring spaces and dashes) = same asset.
- If a serial is missing, the same equipment tag (RTU-4, P-1, LP-2A from the partner's Tag column or the technician's notes) in the same building = same asset. This lets the partner's nameplate fill in units the technician couldn't read.
- Matches on description alone are never merged.
- The same rules catch an asset logged twice in one file ("2nd pass").

**When two records of the same asset disagree.**
- Blanks are filled from the other record. For nameplate fields and location, the partner's values are used because they surveyed nameplates.
- Install year: the technician's install year wins. A gap of 1 to 2 years is expected (the partner records manufacture year). A gap over 2 years goes to review with both values.
- Condition: within 1 point, the worse rating is kept (safer for capital planning). Two or more points apart goes to review with both ratings shown.

**Flag rather than guess.** These go to review:
- Asset types that are vague ("unit", "fan", "heater") or not on FMX's list (Unit Heater, Kitchen hood, Expansion Tank, Plate Heat Exchanger)
- Rows covering several assets ("qty 4")
- Missing nameplates (with the photo file name)
- Unknown buildings or manufacturers
- Notes that change the meaning: abandoned, replaced, locked out, cracked, couldn't reach, come back

## Known limits
- Location text is kept as captured, not standardized.
- A new spelling the rules don't know is flagged, not guessed. Add it to `rules/` and rerun.
- No AI model is called. Every rule is a fixed, explainable lookup.
