import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import pdfplumber
import re
from parser import is_fund_match, get_template_funds_checklist, parse_pdf

def trace():
    pdf_path = r"C:\Users\Darre\OneDrive\Desktop\customer-report-type-D_NG_AIK_SIEW_5205411052-20260606154108.pdf"
    template_path = r"C:\Users\Darre\Downloads\Generated_Report (4).xlsx"
    
    # Run the real parser to get the parsed_funds list
    data = parse_pdf(pdf_path, template_path)
    parsed_fund_names = [f["name"] for f in data["funds"]]
    
    master_checklist = get_template_funds_checklist(template_path)
    for name in parsed_fund_names:
        if name not in master_checklist:
            master_checklist.append(name)
            
    print(f"Master Checklist: {master_checklist}\n")
    
    active_tracking_fund = None
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            page_num = page_idx + 1
            text = page.extract_text() or ""
            lines = text.split('\n')
            
            for line_idx, line in enumerate(lines):
                norm_line = line.strip().lower()
                
                # Check for table header
                is_table_header = "transaction date" in norm_line or "process date" in norm_line or "trx" in norm_line
                
                if not is_table_header:
                    matched_checklist_fund = None
                    for checklist_fund in master_checklist:
                        if is_fund_match(checklist_fund, line):
                            matched_checklist_fund = checklist_fund
                            break
                            
                    if matched_checklist_fund:
                        old_fund = active_tracking_fund
                        active_tracking_fund = matched_checklist_fund
                        if old_fund != active_tracking_fund:
                            print(f"Page {page_num} Line {line_idx+1}: Set active_tracking_fund = '{active_tracking_fund}' (old='{old_fund}') based on line: '{line}'")
                            
                # Check for DIV
                if "DIV" in line:
                    print(f"  --> FOUND DIV on Page {page_num} Line {line_idx+1} under active fund '{active_tracking_fund}': '{line}'")

if __name__ == "__main__":
    trace()
