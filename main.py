import pandas as pd

#1. set file name and sheet name
file_name = "AI_dictionary_consolidated_V1.xlsx"
sheet = "Final Dictionary"

#2. open the excel file and read the sheet
data_table = pd.read_excel(file_name, sheet_name=sheet)

#3. pick out column "Match term"
words_column = data_table["Match term"]

#4. remove empty rows
words_without_blanks = words_column.dropna()

#5. convert the column into python list of words
terms = list(words_without_blanks)

#6. check result
print("Total words found:", len(terms))
print("First 5 words:", terms[0:5])

#8. import tool for reading pdf's
from pypdf import PdfReader
import os

#9. Test with pdf inside the annual reports
sample_pdf_path = os.path.join("Annual reports", "Anima_Holding_2023.pdf")

#10. open pdf
reader = PdfReader(sample_pdf_path)

#10. check how many pages
total_pages = len(reader.pages)
print("Total pages in report:", total_pages)

#11. Extract text from first page to see if works
first_page = reader.pages[0]
first_page_text = first_page.extract_text()

print("Preview of page 1:")
print(first_page_text[:200]) #only first 200 characters

#12. import re for pattern matching
import re

#13. Extract text from all pages into one
full_pdf_text = ""
for page in reader.pages:
    text = page.extract_text()
    if text:
        full_pdf_text += " " + text

#14. Create a dictionary to store term counts
term_counts = {}

#15. Count occurence for each term from Excel list
for term in terms:
    pattern = r"\b" + re.escape(str(term)) + r"\b"    # \b ensures we match complete words, not substrings
    matches = re.findall(pattern, full_pdf_text, re.IGNORECASE)
    term_counts[term] = len(matches)


#16. Show any terms that were found at least once
found_terms = {}

for word, count in term_counts.items():
    if count > 0:
        found_terms[word] = count

print(f"Terms found in this report: {len(found_terms)}")

shown = 0
for word, count in found_terms.items():
    print(f" - {word}: {count}")
    shown += 1
    if shown >= 10:
        break


#17. find all pdfs in annual reports folder
reports_folder = "Annual reports"
all_files = os.listdir(reports_folder)

pdf_files = []
for file in all_files:
    if file.endswith(".pdf"):
        pdf_files.append(file)

print("Found PDF files:", len(pdf_files))
print(pdf_files)