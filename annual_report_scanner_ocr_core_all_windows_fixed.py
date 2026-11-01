"""Scan English annual-report PDFs for AI dictionary matches, with selective OCR.

Install Python packages:
    pip install pandas numpy pypdf pymupdf pytesseract pillow openpyxl

Install the Tesseract OCR application separately; the Windows executable path is configured below.
Windows example: set TESSERACT_CMD to the installed tesseract.exe path below.
"""
from __future__ import annotations

import os
import re
import hashlib
from pathlib import Path
from collections import defaultdict

import numpy as np
import pandas as pd
import pytesseract
from pypdf import PdfReader

# -------- CONFIGURE THESE --------
DICT_FILE = Path("AI_dictionary_consolidated_V1.xlsx")
DICT_SHEET = "Final Dictionary"
PDF_FOLDER = Path(r"Annual reports\2018")  # Change year as needed
OUTPUT_FILE = Path("annual_reports_AI_term_counts_OCR.xlsx")
# If auto-detection fails, set the EXACT Excel header, e.g. "Category":
# Classification is read from "Measure status" (Core / All).
# The descriptive "Final category" labels remain unchanged.
# Optional path if Tesseract is not on Windows PATH:
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
OCR_LANGUAGE = "eng"
OCR_DPI = 175
OCR_SPARSE_PAGE_WORDS = 25
OCR_REPORT_WORDS_PER_PAGE = 100
# Treat reports with fewer words than this as unscannable/invalid, not as zero AI.
MIN_REPORT_WORDS_FOR_VALID_SCORE = 500
# ---------------------------------


def load_dictionary():
    """Use the workbook's explicit Core/All designation, not category numbering.

    'Core' contributes to both Core and All metrics.
    'All' contributes only to the All metric.
    """
    df = pd.read_excel(DICT_FILE, sheet_name=DICT_SHEET)
    df.columns = [str(c).strip() for c in df.columns]
    required = {"Match term", "Final category", "Measure status"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Dictionary missing columns: {sorted(missing)}")

    term_statuses = defaultdict(set)
    term_categories = defaultdict(set)
    for _, row in df.iterrows():
        if pd.isna(row["Match term"]) or not str(row["Match term"]).strip():
            continue
        term = str(row["Match term"]).strip().lower()
        status = str(row["Measure status"]).strip().casefold()
        if status not in {"core", "all"}:
            raise ValueError(f"Unrecognized Measure status for '{term}': {row['Measure status']!r}")
        term_statuses[term].add(status)
        term_categories[term].add(str(row["Final category"]).strip())

    if not term_statuses:
        raise ValueError("No usable dictionary terms found")
    conflicts = [t for t, statuses in term_statuses.items() if len(statuses) > 1]
    if conflicts:
        raise ValueError(f"Conflicting Core/All status on duplicate terms: {conflicts[:10]}")

    terms = sorted(term_statuses, key=lambda t: (-len(t), t))
    patterns = {t: re.compile(r"\b" + re.escape(t) + r"\b", re.I) for t in terms}
    n_core = sum('core' in statuses for statuses in term_statuses.values())
    print(f"Dictionary: {len(terms)} unique terms; Core: {n_core}; All: {len(terms)}")
    return terms, patterns, term_statuses


def word_count(txt):
    return len(re.findall(r"\b\w+\b", txt, re.UNICODE))


def ocr_page(pdf_path, page_index):
    # Imported only when OCR is needed; regular text PDFs work without OCR setup.
    import pymupdf as fitz
    from PIL import Image
    if TESSERACT_CMD:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    with fitz.open(str(pdf_path)) as pdf:
        page = pdf.load_page(page_index)
        pix = page.get_pixmap(dpi=OCR_DPI, alpha=False)
        img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
        return pytesseract.image_to_string(img, lang=OCR_LANGUAGE)


def get_report_text(pdf_path):
    with open(pdf_path, 'rb') as inp:
        reader = PdfReader(inp, strict=False)
        page_texts = []
        errors = 0
        for page in reader.pages:
            try:
                page_texts.append(page.extract_text() or '')
            except Exception as exc:
                print(f"  WARNING: failed to extract a page: {exc}")
                errors += 1
                page_texts.append('')
    n_pages = len(page_texts)
    first_pass_words = sum(word_count(t) for t in page_texts)
    global_sparse = n_pages > 0 and first_pass_words / n_pages < OCR_REPORT_WORDS_PER_PAGE
    ocr_pages = 0
    ocr_failed = 0
    for i, existing in enumerate(page_texts):
        count = word_count(existing)
        # Force OCR on pages having almost no native text; when a report overall
        # is almost unreadable, OCR all its low-word-count pages (<100 words).
        need_ocr = count < OCR_SPARSE_PAGE_WORDS or (global_sparse and count < OCR_REPORT_WORDS_PER_PAGE)
        if not need_ocr:
            continue
        try:
            scanned = ocr_page(pdf_path, i)
            ocr_pages += 1
            # Keep the richer text representation, not both (prevents double counting).
            if word_count(scanned) > count:
                page_texts[i] = scanned
        except Exception as exc:
            ocr_failed += 1
            print(f"  OCR failure on page {i + 1}: {exc}")
            if ocr_failed == 1:
                print("  Check Tesseract installation and TESSERACT_CMD, if required.")
                # Stop repeated costly failures; native extracted text remains usable.
                break

    counts = [word_count(t) for t in page_texts]
    text = '\n'.join(page_texts)
    return text, dict(
        PDF_Pages=n_pages,
        Pages_With_Text=sum(c >= 10 for c in counts),
        Pages_Without_Text=sum(c < 10 for c in counts),
        Text_Coverage_Pct=round(100 * sum(c >= 10 for c in counts) / n_pages, 2) if n_pages else 0,
        Words_Per_Page=round(sum(counts) / n_pages, 2) if n_pages else 0,
        Page_Errors=errors,
        OCR_Pages_Attempted=ocr_pages,
        OCR_Page_Failures=ocr_failed,
    )


def main():
    if not PDF_FOLDER.is_dir():
        raise FileNotFoundError(f"No PDF folder: {PDF_FOLDER}")
    if not DICT_FILE.is_file():
        raise FileNotFoundError(f"Missing dictionary workbook: {DICT_FILE}")
    terms, patterns, term_cats = load_dictionary()
    pdf_files = sorted(p for p in PDF_FOLDER.iterdir() if p.is_file() and p.suffix.lower() == '.pdf')
    print(f"Scanning {len(pdf_files)} PDFs in {PDF_FOLDER}")
    records = []
    for idx, path in enumerate(pdf_files, start=1):
        print(f"[{idx}/{len(pdf_files)}] {path.name}")
        identifier = re.match(r'^(\d+)_', path.name)
        extracted_id = identifier.group(1) if identifier else None
        try:
            doc_text, qc = get_report_text(path)
            total_words = word_count(doc_text)
            valid = total_words >= MIN_REPORT_WORDS_FOR_VALID_SCORE and qc['Page_Errors'] == 0
            term_counts = {term: len(pattern.findall(doc_text)) for term, pattern in patterns.items()}
            all_mentions = sum(term_counts.values())
            core_mentions = sum(term_counts[t] for t in terms if 'core' in term_cats[t])
            status = ('REVIEW — LOW WORD COUNT' if not valid else
                      'REVIEW — PARTIAL OCR' if qc['OCR_Page_Failures'] else
                      'REVIEW — LOW PAGE COVERAGE' if qc['Text_Coverage_Pct'] < 70 else
                      'PASS')
            # Invalid extraction must not be interpreted as "zero AI".
            base = {
                'File_Identifier': extracted_id,
                'Report': path.name,
                'Total_Words': total_words,
                'Total_Mentions': all_mentions if valid else np.nan,
                'AI_Intensity_per_1k': round(all_mentions / total_words * 1000, 4) if valid else np.nan,
                'Log_AI_Mentions': round(float(np.log1p(all_mentions)), 4) if valid else np.nan,
                'Core_Mentions': core_mentions if valid else np.nan,
                'All_Mentions': all_mentions if valid else np.nan,
                'AI_Intensity_Core_per_1k': round(core_mentions / total_words * 1000, 4) if valid else np.nan,
                'AI_Intensity_All_per_1k': round(all_mentions / total_words * 1000, 4) if valid else np.nan,
                'Log_AI_Core_Mentions': round(float(np.log1p(core_mentions)), 4) if valid else np.nan,
                'Log_AI_All_Mentions': round(float(np.log1p(all_mentions)), 4) if valid else np.nan,
                **qc,
                'Extraction_Status': status,
            }
            records.append({**base, **(term_counts if valid else dict.fromkeys(terms, np.nan))})
            print(f"  {total_words} words, OCR attempted on {qc['OCR_Pages_Attempted']} pages; "
                  f"Core={core_mentions if valid else 'NA'}, All={all_mentions if valid else 'NA'}, {status}")
        except Exception as exc:
            print(f"  ERROR: {exc}")
            records.append({'File_Identifier': extracted_id, 'Report': path.name,
                            'Extraction_Status': 'ERROR', 'Extraction_Error': str(exc)})

    columns = ['File_Identifier', 'Report', 'Total_Words', 'Total_Mentions',
               'AI_Intensity_per_1k', 'Log_AI_Mentions', 'Core_Mentions', 'All_Mentions',
               'AI_Intensity_Core_per_1k', 'AI_Intensity_All_per_1k',
               'Log_AI_Core_Mentions', 'Log_AI_All_Mentions', 'PDF_Pages',
               'Pages_With_Text', 'Pages_Without_Text', 'Text_Coverage_Pct',
               'Words_Per_Page', 'Page_Errors', 'OCR_Pages_Attempted',
               'OCR_Page_Failures', 'Extraction_Status', 'Extraction_Error']
    result = pd.DataFrame(records)
    result = result.reindex(columns=columns + terms)
    year_sheet = PDF_FOLDER.name
    # Separate output workbook preserves original results. Replaces only selected year sheet.
    if OUTPUT_FILE.exists():
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            result.to_excel(writer, sheet_name=year_sheet, index=False)
    else:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='w') as writer:
            result.to_excel(writer, sheet_name=year_sheet, index=False)
    print(f"Done: {year_sheet} -> {OUTPUT_FILE} ({len(result)} reports)")


if __name__ == '__main__':
    main()
