import os
import sys
import unittest

# Add backend directory to module search path
backend_dir = os.path.dirname(os.path.dirname(__file__))
if backend_dir not in sys.path:
    sys.path.insert(0, backend_dir)

import auth
from fastapi.testclient import TestClient
from app import app

client = TestClient(app)

class TestAuthAndRevocationSystem(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Reset test DB
        auth.init_db()

    def test_01_admin_login(self):
        print("\n[TEST 1] Testing Admin login with default credentials...")
        res = client.post("/api/auth/login", json={
            "username_or_email": "admin",
            "password": "admin123"
        })
        self.assertEqual(res.status_code, 200, f"Admin login failed: {res.text}")
        data = res.json()
        self.assertIn("token", data)
        self.assertEqual(data["user"]["role"], "admin")
        self.assertEqual(data["user"]["status"], "active")
        print(" -> [PASS] Admin logged in successfully.")

    def test_02_unauthorized_process_rejected(self):
        print("\n[TEST 2] Testing /api/process without auth header (Should be 401)...")
        files = {
            'pdf_file': ('test.pdf', b'%PDF-1.4 dummy', 'application/pdf'),
            'excel_file': ('test.xlsx', b'dummy excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        }
        res = client.post("/api/process", files=files)
        self.assertEqual(res.status_code, 401)
        print(" -> [PASS] Unauthenticated request correctly rejected with 401.")

    def test_03_user_registration_and_revocation_lifecycle(self):
        print("\n[TEST 3] Testing User Registration, Approval, Processing, and Instant Revocation...")
        
        admin_login = client.post("/api/auth/login", json={
            "username_or_email": "admin",
            "password": "admin123"
        })
        admin_token = admin_login.json()["token"]
        admin_headers = {"Authorization": f"Bearer {admin_token}"}
        
        import uuid
        unique_id = str(uuid.uuid4())[:6]
        uname = f"client_{unique_id}"
        uemail = f"client_{unique_id}@fncl.com"

        reg_res = client.post("/api/auth/register", json={
            "full_name": "Test Client",
            "username": uname,
            "email": uemail,
            "password": "secretpassword"
        })
        self.assertEqual(reg_res.status_code, 200)
        print(f" -> [PASS] User '{uname}' registered (status: pending).")
        
        login_pending_res = client.post("/api/auth/login", json={
            "username_or_email": uname,
            "password": "secretpassword"
        })
        self.assertEqual(login_pending_res.status_code, 403)
        self.assertIn("Access Pending", login_pending_res.json()["detail"])
        print(" -> [PASS] Login while pending correctly blocked with 403.")
        
        users_res = client.get("/api/admin/users", headers=admin_headers)
        self.assertEqual(users_res.status_code, 200)
        users = users_res.json()["users"]
        target_user = next(u for u in users if u["username"] == uname)
        
        approve_res = client.post("/api/admin/users/status", json={
            "user_id": target_user["id"],
            "status": "active"
        }, headers=admin_headers)
        self.assertEqual(approve_res.status_code, 200)
        print(f" -> [PASS] Admin approved '{uname}' (status: active).")
        
        login_active_res = client.post("/api/auth/login", json={
            "username_or_email": uname,
            "password": "secretpassword"
        })
        self.assertEqual(login_active_res.status_code, 200)
        client_token = login_active_res.json()["token"]
        client_headers = {"Authorization": f"Bearer {client_token}"}
        print(f" -> [PASS] Approved user '{uname}' logged in successfully.")
        
        revoke_res = client.post("/api/admin/users/status", json={
            "user_id": target_user["id"],
            "status": "revoked"
        }, headers=admin_headers)
        self.assertEqual(revoke_res.status_code, 200)
        print(f" -> [PASS] Admin REVOKED access for '{uname}'!")
        
        me_res = client.get("/api/auth/me", headers=client_headers)
        self.assertEqual(me_res.status_code, 401)
        
        login_revoked_res = client.post("/api/auth/login", json={
            "username_or_email": uname,
            "password": "secretpassword"
        })
        self.assertEqual(login_revoked_res.status_code, 403)
        self.assertIn("Access Revoked", login_revoked_res.json()["detail"])
        print(" -> [PASS] Revoked user token and login attempts immediately blocked with 403!")

if __name__ == "__main__":
    unittest.main()
