import sqlite3
import hashlib
import os
import secrets
from datetime import datetime, timedelta

DB_PATH = os.path.join(os.path.dirname(__file__), "users.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """
    Initializes the SQLite database schema for user authentication and session management.
    """
    with get_db() as conn:
        cursor = conn.cursor()
        
        # Create users table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                username TEXT UNIQUE NOT NULL,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                status TEXT NOT NULL DEFAULT 'pending', -- 'active', 'pending', 'revoked'
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Create sessions table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                token TEXT PRIMARY KEY,
                user_id INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
            )
        """)
        
        conn.commit()
        
        # Seed default Admin account if no admin exists
        cursor.execute("SELECT id FROM users WHERE role = 'admin'")
        if not cursor.fetchone():
            salt, pw_hash = hash_password("admin123")
            cursor.execute("""
                INSERT INTO users (full_name, username, email, password_hash, salt, role, status)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, ("FNCL Administrator", "admin", "admin@fncl.com", pw_hash, salt, "admin", "active"))
            conn.commit()
            print("Initialized default Admin account (username: admin, password: admin123)")

def hash_password(password: str, salt: str = None) -> tuple:
    if not salt:
        salt = secrets.token_hex(16)
    pw_hash = hashlib.pbkdf2_hmac(
        'sha256',
        password.encode('utf-8'),
        salt.encode('utf-8'),
        100000
    ).hex()
    return salt, pw_hash

def verify_password(password: str, salt: str, expected_hash: str) -> bool:
    _, calculated_hash = hash_password(password, salt)
    return secrets.compare_digest(calculated_hash, expected_hash)

def create_user(full_name: str, username: str, email: str, password: str, role: str = "user", status: str = "pending") -> dict:
    username = username.strip().lower()
    email = email.strip().lower()
    full_name = full_name.strip()
    
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM users WHERE username = ? OR email = ?", (username, email))
        if cursor.fetchone():
            raise ValueError("Username or email already exists.")
            
        salt, pw_hash = hash_password(password)
        cursor.execute("""
            INSERT INTO users (full_name, username, email, password_hash, salt, role, status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (full_name, username, email, pw_hash, salt, role, status))
        conn.commit()
        user_id = cursor.lastrowid
        
        return {
            "id": user_id,
            "full_name": full_name,
            "username": username,
            "email": email,
            "role": role,
            "status": status
        }

def authenticate_user(username_or_email: str, password: str) -> dict:
    target = username_or_email.strip().lower()
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ? OR email = ?", (target, target))
        row = cursor.fetchone()
        if not row:
            return None
            
        user = dict(row)
        if not verify_password(password, user["salt"], user["password_hash"]):
            return None
            
        return user

def create_session(user_id: int, days_valid: int = 7) -> str:
    token = secrets.token_urlsafe(32)
    expires_at = datetime.now() + timedelta(days=days_valid)
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO sessions (token, user_id, expires_at)
            VALUES (?, ?, ?)
        """, (token, user_id, expires_at.strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
    return token

def get_user_by_token(token: str) -> dict:
    if not token:
        return None
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT u.id, u.full_name, u.username, u.email, u.role, u.status
            FROM sessions s
            JOIN users u ON s.user_id = u.id
            WHERE s.token = ? AND s.expires_at > ?
        """, (token, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        row = cursor.fetchone()
        if not row:
            return None
        return dict(row)

def delete_session(token: str):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sessions WHERE token = ?", (token,))
        conn.commit()

def list_all_users() -> list:
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, full_name, username, email, role, status, created_at
            FROM users
            ORDER BY id ASC
        """)
        rows = cursor.fetchall()
        return [dict(r) for r in rows]

def set_user_status(user_id: int, new_status: str):
    if new_status not in ("active", "pending", "revoked"):
        raise ValueError("Invalid status value")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET status = ? WHERE id = ?", (new_status, user_id))
        if new_status == "revoked":
            # Invalidate all active sessions for this revoked user immediately
            cursor.execute("DELETE FROM sessions WHERE user_id = ?", (user_id,))
        conn.commit()

def delete_user(user_id: int):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        cursor.execute("DELETE FROM sessions WHERE user_id = ?", (user_id,))
        conn.commit()

def update_user_credentials(user_id: int, new_username: str = None, new_password: str = None):
    with get_db() as conn:
        cursor = conn.cursor()
        if new_username:
            new_username = new_username.strip().lower()
            cursor.execute("SELECT id FROM users WHERE username = ? AND id != ?", (new_username, user_id))
            if cursor.fetchone():
                raise ValueError("Username already taken by another account.")
            cursor.execute("UPDATE users SET username = ? WHERE id = ?", (new_username, user_id))
            
        if new_password:
            salt, pw_hash = hash_password(new_password)
            cursor.execute("UPDATE users SET password_hash = ?, salt = ? WHERE id = ?", (pw_hash, salt, user_id))
            
        conn.commit()

# Initialize DB on module import
init_db()
