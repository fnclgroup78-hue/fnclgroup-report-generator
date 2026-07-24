import urllib.request
import urllib.error
import mimetypes
import os

def build_multipart_formdata(files):
    boundary = b'Boundary-12345'
    body = []
    for field_name, filepath in files.items():
        filename = os.path.basename(filepath)
        mime_type = mimetypes.guess_type(filepath)[0] or 'application/octet-stream'
        body.append(b'--' + boundary)
        body.append(f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"'.encode('utf-8'))
        body.append(f'Content-Type: {mime_type}'.encode('utf-8'))
        body.append(b'')
        with open(filepath, 'rb') as f:
            body.append(f.read())
    body.append(b'--' + boundary + b'--')
    body.append(b'')
    content_type = f'multipart/form-data; boundary={boundary.decode("utf-8")}'
    return content_type, b'\r\n'.join(body)

def test_api():
    print("Running API Validation checks...")
    
    # 1. Prepare files
    pdf_path = "tests/manulife_statement.pdf"
    xlsx_path = "tests/template.xlsx"
    xls_path = "tests/test_invalid.xls"
    
    # Create a dummy .xls file
    with open(xls_path, "wb") as f:
        f.write(b"dummy excel contents")
        
    try:
        # Test Case 1: Uploading .xls should fail
        print("Testing Case 1: Uploading .xls file (Should fail)...")
        content_type, body = build_multipart_formdata({
            'pdf_file': pdf_path,
            'excel_file': xls_path
        })
        req = urllib.request.Request(
            'http://127.0.0.1:8080/api/process',
            data=body,
            headers={'Content-Type': content_type}
        )
        try:
            with urllib.request.urlopen(req) as res:
                print("FAIL: Uploading .xls file succeeded, but should have failed!")
                exit(1)
        except urllib.error.HTTPError as e:
            assert e.code == 400, f"Expected 400 status, got {e.code}"
            error_msg = e.read().decode('utf-8')
            assert "must be an Excel spreadsheet (.xlsx)" in error_msg, f"Unexpected error message: {error_msg}"
            print("[PASS] Case 1: .xls upload correctly rejected with 400 and detail message.")
            
        # Test Case 2: Uploading correct .xlsx should succeed
        print("Testing Case 2: Uploading correct .xlsx file (Should succeed)...")
        content_type, body = build_multipart_formdata({
            'pdf_file': pdf_path,
            'excel_file': xls_path
        })
        # Let's fix content to use xlsx
        content_type, body = build_multipart_formdata({
            'pdf_file': pdf_path,
            'excel_file': xls_path # wait, this was xls_path. Let's do xlsx_path:
        })
        content_type, body = build_multipart_formdata({
            'pdf_file': pdf_path,
            'excel_file': xlsx_path
        })
        req = urllib.request.Request(
            'http://127.0.0.1:8080/api/process',
            data=body,
            headers={'Content-Type': content_type}
        )
        with urllib.request.urlopen(req) as res:
            assert res.status == 200, f"Expected 200 status, got {res.status}"
            assert res.headers.get_content_type() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Unexpected content type"
            data = res.read()
            assert len(data) > 0, "Response is empty"
            print("[PASS] Case 2: Valid files processed successfully.")
            
    finally:
        # Cleanup
        if os.path.exists(xls_path):
            os.remove(xls_path)
            
    print("\nAPI VALIDATION SUCCESSFUL: 100% CORRECT!")

if __name__ == "__main__":
    test_api()
