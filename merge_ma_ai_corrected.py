"""Merge M&A master with annual-report AI scores without filtering on controls.

Run from the folder containing the input files:
    python merge_ma_ai.py

Only M&A transactions with a uniquely matched, valid AI score are exported to
'Matched_MA_AI'. Genuine zero scores are retained. Missing M&A fields remain blank.
"""
from pathlib import Path
import pandas as pd

# Adjust paths if your actual files have different names/locations.
MA_FILE = Path("MA_master_final.xlsx")
AI_FILE = Path("annual_reports_AI_term_counts_OCR.xlsx")
OUTPUT_FILE = Path("MA_AI_merged_corrected.xlsx")
MA_SHEET = "Master Data"


def clean_id(value):
    """Convert Excel numeric/text identifiers to a consistent string."""
    if pd.isna(value):
        return None
    value = str(value).strip()
    if not value or value.lower() in ("nan", "none", "null"):
        return None
    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]
    return value


def main():
    if not MA_FILE.exists() or not AI_FILE.exists():
        missing = [str(p) for p in (MA_FILE, AI_FILE) if not p.exists()]
        raise FileNotFoundError("Missing file(s): " + ", ".join(missing))

    ma = pd.read_excel(MA_FILE, sheet_name=MA_SHEET, dtype={
        "SDC Deal No": "string", "Target PermID": "string"
    })
    ma.columns = ma.columns.astype(str).str.strip()
    required_ma = {"SDC Deal No", "Target PermID", "filing_year"}
    if not required_ma.issubset(ma.columns):
        raise ValueError("M&A file missing columns: " + str(required_ma - set(ma.columns)))

    xls = pd.ExcelFile(AI_FILE)
    years = [s for s in xls.sheet_names if s.isdigit() and len(s) == 4]
    if not years:
        raise ValueError("No fiscal-year worksheets (e.g. 2015) found in AI workbook.")
    reports = []
    for year in sorted(years):
        df = pd.read_excel(xls, sheet_name=year, dtype={"File_Identifier": "string"})
        df["AI_Report_FY"] = int(year)
        reports.append(df)
    ai = pd.concat(reports, ignore_index=True)
    required_ai = {"File_Identifier", "AI_Intensity_Core_per_1k", "AI_Intensity_All_per_1k"}
    if not required_ai.issubset(ai.columns):
        raise ValueError("AI workbook missing columns: " + str(required_ai - set(ai.columns)))

    ai["_id"] = ai["File_Identifier"].map(clean_id)
    ai["AI_Report_FY"] = pd.to_numeric(ai["AI_Report_FY"], errors="coerce").astype("Int64")
    ai["AI_Intensity_Core_per_1k"] = pd.to_numeric(ai["AI_Intensity_Core_per_1k"], errors="coerce")
    ai["AI_Intensity_All_per_1k"] = pd.to_numeric(ai["AI_Intensity_All_per_1k"], errors="coerce")
    ai_valid = ai.loc[
        ai["_id"].notna()
        & ai["AI_Intensity_Core_per_1k"].notna()
        & ai["AI_Intensity_All_per_1k"].notna()
        & ai["Total_Words"].gt(0)
        & ai["Extraction_Status"].eq("PASS")
    ].copy()

    # Build a crosswalk from ANY known transaction SDC number to its target PermID.
    # This lets an FY2017 report named after an older SDC number match an FY2017
    # transaction with a different SDC number but the same target PermID.
    sdc_to_permids = {}
    for _, tx in ma.iterrows():
        sdc = clean_id(tx["SDC Deal No"])
        pid = clean_id(tx["Target PermID"])
        if sdc and pid:
            sdc_to_permids.setdefault(sdc, set()).add(pid)

    # Index by fiscal year + identifier. Several PDFs can share the same numeric
    # ID; ambiguous keys are NOT arbitrarily assigned to a transaction.
    lookup = {}
    for row_id, row in ai_valid.iterrows():
        lookup.setdefault((int(row["AI_Report_FY"]), row["_id"]), []).append(row_id)

    # Index AI reports by target PermID derived from their filename identifier.
    # Direct PermID filenames also map to themselves. Ambiguous mappings are
    # deliberately excluded from automatic crosswalk matching.
    permid_lookup = {}
    all_permids = set(ma["Target PermID"].map(clean_id).dropna())
    for row_id, row in ai_valid.iterrows():
        identifier = row["_id"]
        pids = sdc_to_permids.get(identifier, set())
        if identifier in all_permids:
            pids = pids | {identifier}
        if len(pids) == 1:
            pid = next(iter(pids))
            permid_lookup.setdefault((int(row["AI_Report_FY"]), pid), []).append(row_id)

    ma["_sdc"] = ma["SDC Deal No"].map(clean_id)
    ma["_permid"] = ma["Target PermID"].map(clean_id)
    ma["_fy"] = pd.to_numeric(ma["filing_year"], errors="coerce")

    matched_pairs = []
    audit_rows = []
    for i, tx in ma.iterrows():
        fy = int(tx["_fy"]) if pd.notna(tx["_fy"]) else None
        sdc_hits = lookup.get((fy, tx["_sdc"]), []) if fy and tx["_sdc"] else []
        pid_hits = permid_lookup.get((fy, tx["_permid"]), []) if fy and tx["_permid"] else []
        candidates = list(dict.fromkeys(sdc_hits + pid_hits))
        if len(candidates) == 1:
            aidx = candidates[0]
            if aidx in sdc_hits:
                method = "SDC Deal No"
            elif ai_valid.loc[aidx, "_id"] == tx["_permid"]:
                method = "Target PermID"
            else:
                method = "SDC-to-PermID crosswalk"
            matched_pairs.append((i, aidx, method))
            status = "MATCHED"
        elif len(candidates) > 1:
            status, method = "AMBIGUOUS_MATCH", ""
        else:
            status, method = "NO_VALID_AI_MATCH", ""
        audit_rows.append({
            "M&A_row": i + 2,
            "SDC Deal No": tx["SDC Deal No"],
            "Target PermID": tx["Target PermID"],
            "Target Full Name": tx.get("Target Full Name"),
            "filing_year": tx["filing_year"],
            "Merge_Status": status,
            "Match_Method": method,
            "Candidate_AI_Reports": len(candidates),
        })

    # Keep every original M&A column (including missing premiums/controls).
    ma_original_cols = [c for c in ma.columns if not c.startswith("_")]
    ai_export_cols = [c for c in ai_valid.columns if not c.startswith("_")]
    merged_rows = []
    for midx, aidx, method in matched_pairs:
        record = ma.loc[midx, ma_original_cols].to_dict()
        record["Merge_Method"] = method
        for col in ai_export_cols:
            record["AI_" + col] = ai_valid.loc[aidx, col]
        merged_rows.append(record)

    merged = pd.DataFrame(merged_rows)
    audit = pd.DataFrame(audit_rows)
    summary = pd.DataFrame([
        {"Metric": "M&A transactions in master", "Value": len(ma)},
        {"Metric": "Valid annual-report observations", "Value": len(ai_valid)},
        {"Metric": "Matched transactions", "Value": len(merged)},
        {"Metric": "Unmatched transactions", "Value": int((audit.Merge_Status == "NO_VALID_AI_MATCH").sum())},
        {"Metric": "Ambiguous transactions", "Value": int((audit.Merge_Status == "AMBIGUOUS_MATCH").sum())},
    ])

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Matched_MA_AI", index=False)
        audit.to_excel(writer, sheet_name="Merge_Audit", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"Merged {len(merged)} of {len(ma)} transactions.")
    print(f"Ambiguous matches: {(audit.Merge_Status == 'AMBIGUOUS_MATCH').sum()}")
    print(f"Crosswalk matches: {(merged.Merge_Method == 'SDC-to-PermID crosswalk').sum()}")
    print(f"Saved: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
