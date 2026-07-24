import pdfplumber
import os

pdf_path = r"C:\Users\Darre\OneDrive\Desktop\customer-report-type-D_NG_AIK_SIEW_5205411052-20260606154108.pdf"

with pdfplumber.open(pdf_path) as pdf:
    for page_idx, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if "DIV" in line:
                print(f"--- Page {page_idx + 1}, Line {i + 1} ---")
                start = max(0, i - 4)
                end = min(len(lines), i + 5)
                for j in range(start, end):
                    marker = ">>> " if j == i else "    "
                    print(f"{marker}Line {j+1}: {lines[j]}")
