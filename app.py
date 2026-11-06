"""
Web page for the FMX reconciliation tool.
Run locally:  streamlit run app.py
It uses the same logic as reconcile.py; this file only adds upload, run, and download.
"""

import tempfile
from argparse import Namespace
from pathlib import Path

import pandas as pd
import streamlit as st

import reconcile

HERE = Path(__file__).parent
DATA = HERE / "data"

st.set_page_config(page_title="FCA Reconciliation", layout="wide")
st.title("FCA Reconciliation")
st.write("Upload the technician capture file and the partner equipment list. "
         "The tool fixes values with a clear right answer and flags the rest for review.")

col1, col2 = st.columns(2)
tech_file = col1.file_uploader("Technician capture (.xlsx or .csv)", type=["xlsx", "csv"])
partner_file = col2.file_uploader("Partner equipment list (optional)", type=["xlsx", "csv"])

with st.expander("Reference files (FMX asset types and import format)"):
    st.write("By default the tool uses the copies in the data folder. Upload new ones only if FMX's lists changed.")
    types_file = st.file_uploader("Asset types", type=["xlsx", "csv"])
    target_file = st.file_uploader("Target format", type=["xlsx", "csv"])


def save(upload, folder):
    path = Path(folder) / upload.name
    path.write_bytes(upload.getvalue())
    return path


if st.button("Run", type="primary", disabled=tech_file is None):
    with tempfile.TemporaryDirectory() as tmp:
        args = Namespace(
            tech=save(tech_file, tmp),
            partner=save(partner_file, tmp) if partner_file else "",
            asset_types=save(types_file, tmp) if types_file else DATA / "Asset_Types.xlsx",
            target=save(target_file, tmp) if target_file else DATA / "Target_Format.xlsx",
            out=Path(tmp) / "output",
        )
        try:
            reconcile.run(args)
        except SystemExit as e:  # the tool stops with a plain message when a file can't be read
            st.error(str(e))
            st.stop()

        out = args.out
        st.session_state["results"] = {
            "summary": (out / "run_summary.txt").read_text(),
            "import": (out / "import_ready.csv").read_bytes(),
            "review": (out / "review_list.csv").read_bytes(),
        }

results = st.session_state.get("results")
if results:
    imp = pd.read_csv(pd.io.common.BytesIO(results["import"]), dtype=str).fillna("")
    rev = pd.read_csv(pd.io.common.BytesIO(results["review"]), dtype=str).fillna("")

    st.subheader("Run summary")
    c = st.columns(4)
    c[0].metric("Rows out", len(imp))
    c[1].metric("Ready", int((imp["status"] == "ready").sum()))
    c[2].metric("Needs review", int((imp["status"] == "needs_review").sum()))
    merged = next((l for l in results["summary"].splitlines() if l.startswith("Duplicates merged")), "")
    c[3].metric("Duplicates merged", merged.split(":")[1].split("(")[0].strip() if merged else "0")
    with st.expander("Full summary"):
        st.code(results["summary"], language=None)

    st.subheader("Rows that need a person")
    if len(rev):
        first = ["review_reason", "building", "asset_type", "original_asset_type", "manufacturer",
                 "model", "serial_number", "install_year", "condition", "location", "source", "source_rows"]
        rev = rev[[c for c in first if c in rev.columns] + [c for c in rev.columns if c not in first]]
    st.dataframe(rev, width="stretch", hide_index=True)
    st.subheader("Download")
    d = st.columns(3)
    d[0].download_button("import_ready.csv", results["import"], "import_ready.csv", "text/csv")
    d[1].download_button("review_list.csv", results["review"], "review_list.csv", "text/csv")
    d[2].download_button("run_summary.txt", results["summary"], "run_summary.txt", "text/plain")
