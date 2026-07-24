import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from parser import parse_pdf
import pprint

def test_ng():
    pdf_path = r"C:\Users\Darre\OneDrive\Desktop\customer-report-type-D_NG_AIK_SIEW_5205411052-20260606154108.pdf"
    excel_template = r"C:\Users\Darre\Downloads\Generated_Report (4).xlsx"
    
    print("Parsing PDF Statement...")
    data = parse_pdf(pdf_path, excel_template)
    
    print("\nParsed metadata:")
    print(f"Statement Period: {data['statement_period']}")
    print(f"Account Holder: {data['account_holder']}")
    print(f"Joint Holder: {data['joint_holder']}")
    print(f"Account No: {data['account_no']}")
    print(f"Investment Type: {data['investment_type']}")
    
    print("\nParsed Funds:")
    pprint.pprint(data['funds'])
    
    print("\nParsed Distributions:")
    pprint.pprint(data['distributions'])

if __name__ == "__main__":
    test_ng()
