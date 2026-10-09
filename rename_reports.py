import os
import re
import pandas as pd

# Paths
folder_path = r"C:\Users\crist\Desktop\Kouluhommat\Gradu\AI_dicitonary_Construction\Annual reports\2023"
master_excel_path = (
    r"C:\Users\crist\Desktop\Kouluhommat\Gradu\Data\Final data\MA_master_final.xlsx"
)

# 1. Load Master Excel
try:
  df = pd.read_excel(master_excel_path, sheet_name="Master Data")
except Exception:
  df = pd.read_excel(master_excel_path)

deal_col = [c for c in df.columns if "deal" in c.lower()][0]
name_col = [
    c
    for c in df.columns
    if "target" in c.lower() and ("name" in c.lower() or "full" in c.lower())
][0]
ticker_col = [
    c
    for c in df.columns
    if any(k in c.lower() for k in ["ticker", "symbol", "ric"])
][0]

df[deal_col] = df[deal_col].astype(str).str.split(".").str[0]

# Known aliases for LSEG naming quirks
ALIASES = {
    "HELIOS RE SOCIMI": "Lar Espana",
    "ACTIUM HOLDINGS": "Alpha Financial",
}

renamed_count = 0
unmatched = []

for filename in os.listdir(folder_path):
  if not filename.lower().endswith(".pdf"):
    continue

  # Skip files that have already been renamed (starts with deal number)
  if re.match(r"^\d{9,12}_", filename):
    continue

  parts = [p.strip() for p in filename.split(" - ")]
  if len(parts) < 4:
    unmatched.append((filename, "Filename structure mismatch"))
    continue

  file_ticker = parts[2].split(".")[0].upper()
  file_comp_name = parts[3].upper()

  # Apply alias if exists
  search_name = file_comp_name
  for k, v in ALIASES.items():
    if k in file_comp_name:
      search_name = v.upper()
      break

  # Step A: Match by Ticker
  match = df[
      df[ticker_col]
      .astype(str)
      .str.upper()
      .str.contains(file_ticker, regex=False, na=False)
  ]

  # Step B: Match by Company Name (first 8 characters) if ticker fails
  if match.empty:
    core_name = re.sub(r"[^\w]", "", search_name)[:8]
    match = df[
        df[name_col]
        .astype(str)
        .str.upper()
        .str.replace(r"[^\w]", "", regex=True)
        .str.contains(core_name, na=False)
    ]

  # Step C: Fallback check against title if still empty
  if match.empty and len(parts) >= 6:
    file_title = parts[5].upper()
    for _, row in df.iterrows():
      t_clean = re.sub(r"[^\w]", "", str(row[name_col]).upper())[:8]
      if t_clean in re.sub(r"[^\w]", "", file_title):
        match = df[df[deal_col] == row[deal_col]]
        break

  if not match.empty:
    deal_no = match.iloc[0][deal_col]
    target_clean = re.sub(r"[^\w\-_]", "_", str(match.iloc[0][name_col]))[:25]
    new_filename = f"{deal_no}_{target_clean}.pdf"

    old_file_path = os.path.join(folder_path, filename)
    new_file_path = os.path.join(folder_path, new_filename)

    os.rename(old_file_path, new_file_path)
    print(f"Renamed: {filename[:30]}... -> {new_filename}")
    renamed_count += 1
  else:
    unmatched.append((filename, f"Could not match: {file_comp_name}"))

print(f"\nCompleted! Renamed: {renamed_count} files.")
if unmatched:
  print(f"\nUnmatched files ({len(unmatched)}):")
  for f, r in unmatched:
    print(f" - {f}: {r}")