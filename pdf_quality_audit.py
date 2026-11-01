
import os
import re
import numpy as np
import pandas as pd
from pypdf import PdfReader

# 1. Paths
dict_file = "AI_dictionary_consolidated_V1.xlsx"
dict_sheet = "Final Dictionary"

pdf_folder = r"Annual reports\2024"
output_file = "annual_reports_AI_term_counts.xlsx"

sheet_name = os.path.basename(os.path.normpath(pdf_folder))

# 2. Load dictionary
dict_df = pd.read_excel(dict_file, sheet_name=dict_sheet)

terms = (
    dict_df["Match term"]
    .dropna()
    .astype(str)
    .str.strip()
    .str.lower()
    .unique()
    .tolist()
)

terms = sorted(terms, key=len, reverse=True)

patterns = {
    term: re.compile(r"\b" + re.escape(term) + r"\b")
    for term in terms
}

print(f"Loaded {len(terms)} unique search terms.")

# 3. Find PDF files
if not os.path.isdir(pdf_folder):
    raise FileNotFoundError(pdf_folder)

pdf_files = sorted(
    f for f in os.listdir(pdf_folder)
    if f.lower().endswith(".pdf")
)

print(f"Found {len(pdf_files)} PDFs.")

results = []

# 4. Scan reports
for file in pdf_files:
    path = os.path.join(pdf_folder, file)

    print(f"Analyzing {file}...")

    identifier_match = re.match(r"^(\d+)_", file)
    identifier = (
        identifier_match.group(1)
        if identifier_match else None
    )

    page_texts = []
    page_word_counts = []
    pages_without_text = 0
    page_errors = 0

    try:
        reader = PdfReader(path)
        pdf_pages = len(reader.pages)

        for page in reader.pages:
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
                page_errors += 1

            page_texts.append(text)

            word_count = len(re.findall(r"\b\w+\b", text))
            page_word_counts.append(word_count)

            # A page with fewer than 10 extracted
            # words is treated as text-sparse.
            if word_count < 10:
                pages_without_text += 1

    except Exception as e:
        print(f"PDF read error: {e}")
        continue

    doc_text = " ".join(page_texts).lower()
    total_words = sum(page_word_counts)

    pages_with_text = pdf_pages - pages_without_text

    coverage_pct = (
        100 * pages_with_text / pdf_pages
        if pdf_pages else 0
    )

    words_per_page = (
        total_words / pdf_pages
        if pdf_pages else 0
    )

    # 5. Extraction quality classification
    # These thresholds are screening rules,
    # not proof that extraction is complete.
    if total_words == 0:
        extraction_status = "OCR REQUIRED"

    elif (
        page_errors > 0
        or total_words < 5000
        or coverage_pct < 70
        or words_per_page < 100
    ):
        extraction_status = "REVIEW"

    else:
        extraction_status = "PASS"

    # 6. Dictionary matching
    counts = {}

    for term, pattern in patterns.items():
        counts[term] = len(pattern.findall(doc_text))

    total_mentions = sum(counts.values())

    # Missing extraction is not a true zero.
    if total_words > 0:
        ai_intensity = total_mentions / total_words * 1000
        log_mentions = np.log1p(total_mentions)
    else:
        ai_intensity = np.nan
        log_mentions = np.nan

    # 7. Output row
    row = {
        "File_Identifier": identifier,
        "Report": file,
        "Total_Words": total_words,
        "Total_Mentions": total_mentions,
        "AI_Intensity_per_1k": (
            round(ai_intensity, 4)
            if pd.notna(ai_intensity) else np.nan
        ),
        "Log_AI_Mentions": (
            round(log_mentions, 4)
            if pd.notna(log_mentions) else np.nan
        ),
        "PDF_Pages": pdf_pages,
        "Pages_With_Text": pages_with_text,
        "Pages_Without_Text": pages_without_text,
        "Text_Coverage_Pct": round(coverage_pct, 2),
        "Words_Per_Page": round(words_per_page, 2),
        "Page_Errors": page_errors,
        "Extraction_Status": extraction_status,
    }

    row.update(counts)
    results.append(row)

# 8. Save to fiscal-year worksheet
output_df = pd.DataFrame(results)

# Keep the familiar scanner columns first.
base_cols = [
    "File_Identifier",
    "Report",
    "Total_Words",
    "Total_Mentions",
    "AI_Intensity_per_1k",
    "Log_AI_Mentions",
    "PDF_Pages",
    "Pages_With_Text",
    "Pages_Without_Text",
    "Text_Coverage_Pct",
    "Words_Per_Page",
    "Page_Errors",
    "Extraction_Status",
]

other_cols = [
    col for col in output_df.columns
    if col not in base_cols
]

output_df = output_df[base_cols + other_cols]

if os.path.exists(output_file):
    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        output_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )
else:
    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
        mode="w"
    ) as writer:
        output_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

# 9. Print audit summary
print("\nExtraction audit:")
print(output_df["Extraction_Status"].value_counts())

flagged = output_df[
    output_df["Extraction_Status"] != "PASS"
]

print("\nReports requiring attention:")
print(
    flagged[
        ["Report", "Total_Words",
         "Text_Coverage_Pct", "Extraction_Status"]
    ].to_string(index=False)
)

print(f"\nSaved worksheet: {sheet_name}")
