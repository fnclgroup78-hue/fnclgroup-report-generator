import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import pdfplumber
import re
from parser import is_fund_match, get_template_funds_checklist, parse_pdf, normalize_fund_name

pdf_path = r"C:\Users\Darre\OneDrive\Desktop\customer-report-type-D_NG_AIK_SIEW_5205411052-20260606154108.pdf"
template_path = r"C:\Users\Darre\Downloads\Generated_Report (4).xlsx"

def debug_matching():
    # Load checklist
    master_checklist = get_template_funds_checklist(template_path)
    
    # Run the real parser to get the parsed_funds list
    data = parse_pdf(pdf_path, template_path)
    parsed_fund_names = [f["name"] for f in data["funds"]]
    
    for name in parsed_fund_names:
        if name not in master_checklist:
            master_checklist.append(name)
            
    # Add fallback defaults
    default_fallbacks = [
        "Manulife Asia-Pacific REIT Fund",
        "Manulife Global Managed Fund",
        "Manulife Investment Shariah Progress-Plus Fund",
        "Manulife Cash Management Fund",
        "SHARIAH ASIA-PACIFIC EX JAPAN",
        "SHARIAH PROGRESS",
        "ASIA-PACIFIC REIT"
    ]
    for name in default_fallbacks:
        if not any(is_fund_match(name, x) for x in master_checklist):
            master_checklist.append(name)
            
    master_checklist = sorted(master_checklist, key=lambda x: len(x), reverse=True)
    
    print("Master checklist sorted:")
    for c in master_checklist:
        print(f" - '{c}'")
        
    print("\nTracing line-by-line matches on Pages 2 to 6:")
    
    with pdfplumber.open(pdf_path) as pdf:
        for idx in [1, 2, 3, 4, 5]: # pages 2 to 6 (0-indexed: 1, 2, 3, 4, 5)
            page_num = idx + 1
            print(f"\n--- PAGE {page_num} ---")
            page_obj = pdf.pages[idx]
            words = page_obj.extract_words() or []
            words = sorted(words, key=lambda w: (w["top"], w["x0"]))
            
            page_lines = []
            curr_line_words = []
            last_top = None
            for w in words:
                if last_top is None or abs(w["top"] - last_top) < 3:
                    curr_line_words.append(w)
                else:
                    line_text = " ".join([x["text"] for x in curr_line_words])
                    cleaned_line = line_text.replace('\n', ' ').replace('\r', ' ')
                    cleaned_line = re.sub(r'\s+', ' ', cleaned_line).strip()
                    page_lines.append(cleaned_line)
                    curr_line_words = [w]
                last_top = w["top"]
            if curr_line_words:
                line_text = " ".join([x["text"] for x in curr_line_words])
                cleaned_line = line_text.replace('\n', ' ').replace('\r', ' ')
                cleaned_line = re.sub(r'\s+', ' ', cleaned_line).strip()
                page_lines.append(cleaned_line)
                
            # Print all lines and if they match anything in checklist
            for line_idx, line in enumerate(page_lines):
                norm_line = line.strip().lower()
                is_table_header = "transaction date" in norm_line or "process date" in norm_line or "trx" in norm_line
                
                matched = []
                for checklist_fund in master_checklist:
                    if is_fund_match(checklist_fund, line):
                        matched.append(checklist_fund)
                        
                if matched or "DIV" in line:
                    print(f"Line {line_idx+1:02d}: '{line}'")
                    print(f"   Matches: {matched} | TableHeader: {is_table_header}")
                    if "DIV" in line:
                        print(f"   --> HAS DIV")

if __name__ == "__main__":
    debug_matching()
