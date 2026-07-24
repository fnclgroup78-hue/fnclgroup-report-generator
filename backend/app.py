from fastapi import FastAPI, Request, Response, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.exceptions import RequestValidationError
from starlette.exceptions import HTTPException as StarletteHTTPException
from pydantic import BaseModel
from typing import Optional
import os
import shutil
import uuid
import traceback
from parser import parse_pdf
from excel_writer import update_excel_report
import auth

app = FastAPI(title="FNCL Group Manulife Investment Report Generator")

# Pydantic Schemas for Auth and Admin
class RegisterRequest(BaseModel):
    full_name: str
    username: str
    email: str
    password: str

class LoginRequest(BaseModel):
    username_or_email: str
    password: str

class StatusUpdateRequest(BaseModel):
    user_id: int
    status: str

class CreateUserAdminRequest(BaseModel):
    full_name: str
    username: str
    email: str
    password: str
    role: Optional[str] = "user"
    status: Optional[str] = "active"

class ChangeCredentialsRequest(BaseModel):
    current_password: str
    new_username: Optional[str] = None
    new_password: Optional[str] = None

# Auth Dependency Helper
def get_current_user(request: Request) -> dict:
    auth_header = request.headers.get("Authorization")
    token = None
    if auth_header and auth_header.startswith("Bearer "):
        token = auth_header.split(" ", 1)[1].strip()
    elif "x-auth-token" in request.headers:
        token = request.headers.get("x-auth-token").strip()
        
    if not token:
        raise HTTPException(status_code=401, detail="Authentication token required. Please sign in.")
        
    user = auth.get_user_by_token(token)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid or expired session. Please sign in again.")
        
    if user["status"] == "revoked":
        raise HTTPException(status_code=403, detail="Access Revoked: Your account access has been revoked by FNCL Group Administrator.")
    elif user["status"] == "pending":
        raise HTTPException(status_code=403, detail="Access Pending: Your account is awaiting Administrator approval.")
        
    return user


# Custom CORS and Private Network Access (PNA) Middleware
@app.middleware("http")
async def cors_and_pna_middleware(request: Request, call_next):
    if request.method == "OPTIONS":
        response = Response(status_code=200)
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE, PATCH"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Requested-With, x-auth-token"
        response.headers["Access-Control-Allow-Private-Network"] = "true"
        return response

    try:
        response = await call_next(request)
    except Exception as e:
        print(f"Error handling request {request.method} {request.url.path}: {e}")
        traceback.print_exc()
        response = JSONResponse(
            status_code=500,
            content={"detail": f"Internal Server Error: {str(e)}"}
        )

    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE, PATCH"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Requested-With, x-auth-token"
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    
    return response


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    print(f"Validation Error at {request.url.path}: {exc.errors()}")
    error_details = []
    for err in exc.errors():
        loc_path = " -> ".join(str(x) for x in err.get("loc", []))
        msg = err.get("msg", "invalid value")
        error_details.append(f"{loc_path}: {msg}")
    friendly_msg = "; ".join(error_details)
    
    return JSONResponse(
        status_code=422,
        content={"detail": f"Request validation failed: {friendly_msg}"}
    )


@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(request: Request, exc: StarletteHTTPException):
    print(f"HTTP Exception at {request.url.path}: status={exc.status_code}, detail={exc.detail}")
    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": exc.detail}
    )


@app.get("/api/health")
async def health_check():
    return {"status": "ok"}


@app.get("/api/template/download")
async def download_blank_template(request: Request, type: Optional[str] = "main"):
    current_user = get_current_user(request)
    filename = "FNCL ML MAIN REPORT.xlsx" if type != "simple" else "FNCL ML MAIN REPORT (SImple).xlsx"
    
    project_root = os.path.dirname(os.path.dirname(__file__))
    template_path = os.path.join(project_root, filename)
    
    if not os.path.exists(template_path):
        template_path = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(template_path):
            raise HTTPException(status_code=404, detail=f"Excel template '{filename}' not found on server.")
            
    return FileResponse(
        path=template_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# AUTHENTICATION ENDPOINTS
@app.post("/api/auth/register")
async def register(req: RegisterRequest):
    try:
        user = auth.create_user(
            full_name=req.full_name,
            username=req.username,
            email=req.email,
            password=req.password,
            role="user",
            status="pending"
        )
        return {"status": "ok", "message": "Registration successful! Account pending Administrator approval.", "user": user}
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Registration failed: {str(e)}")


@app.post("/api/auth/login")
async def login(req: LoginRequest):
    user = auth.authenticate_user(req.username_or_email, req.password)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid username or password.")
        
    if user["status"] == "revoked":
        raise HTTPException(status_code=403, detail="Access Revoked: Your account access has been revoked by FNCL Group Administrator.")
    elif user["status"] == "pending":
        raise HTTPException(status_code=403, detail="Access Pending: Your account is awaiting Administrator approval.")
        
    token = auth.create_session(user["id"])
    return {
        "token": token,
        "user": {
            "id": user["id"],
            "full_name": user["full_name"],
            "username": user["username"],
            "email": user["email"],
            "role": user["role"],
            "status": user["status"]
        }
    }


@app.get("/api/auth/me")
async def get_me(request: Request):
    user = get_current_user(request)
    return {"user": user}


@app.post("/api/auth/logout")
async def logout(request: Request):
    auth_header = request.headers.get("Authorization")
    if auth_header and auth_header.startswith("Bearer "):
        token = auth_header.split(" ", 1)[1].strip()
        auth.delete_session(token)
    elif "x-auth-token" in request.headers:
        auth.delete_session(request.headers["x-auth-token"])
    return {"status": "ok"}


@app.post("/api/auth/change-credentials")
async def change_credentials(req: ChangeCredentialsRequest, request: Request):
    user = get_current_user(request)
    
    db_user = auth.authenticate_user(user["username"], req.current_password)
    if not db_user:
        raise HTTPException(status_code=400, detail="Incorrect current password.")
        
    if not req.new_username and not req.new_password:
        raise HTTPException(status_code=400, detail="Please enter a new username or a new password to update.")
        
    try:
        auth.update_user_credentials(user["id"], new_username=req.new_username, new_password=req.new_password)
        return {"status": "ok", "message": "Credentials updated successfully!"}
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))


# ADMIN MANAGEMENT ENDPOINTS
@app.get("/api/admin/users")
async def admin_list_users(request: Request):
    user = get_current_user(request)
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin privileges required.")
    users = auth.list_all_users()
    return {"users": users}


@app.post("/api/admin/users/create")
async def admin_create_user(req: CreateUserAdminRequest, request: Request):
    user = get_current_user(request)
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin privileges required.")
    try:
        new_user = auth.create_user(
            full_name=req.full_name,
            username=req.username,
            email=req.email,
            password=req.password,
            role=req.role or "user",
            status=req.status or "active"
        )
        return {"status": "ok", "user": new_user}
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))


@app.post("/api/admin/users/status")
async def admin_update_user_status(req: StatusUpdateRequest, request: Request):
    user = get_current_user(request)
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin privileges required.")
    try:
        auth.set_user_status(req.user_id, req.status)
        return {"status": "ok", "message": f"User status updated to {req.status}"}
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))


@app.delete("/api/admin/users/{user_id}")
async def admin_delete_user(user_id: int, request: Request):
    user = get_current_user(request)
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin privileges required.")
    if user["id"] == user_id:
        raise HTTPException(status_code=400, detail="Cannot delete your own administrator account.")
    auth.delete_user(user_id)
    return {"status": "ok", "message": "User deleted"}


# Ensure temporary upload directories exist
TEMP_DIR = os.path.join(os.path.dirname(__file__), "temp_uploads")
os.makedirs(TEMP_DIR, exist_ok=True)

def cleanup_temp_files(*filepaths):
    for path in filepaths:
        try:
            if os.path.exists(path):
                os.remove(path)
                print(f"Cleaned up temp file: {path}")
        except Exception as e:
            print(f"Error cleaning up temp file {path}: {e}")


@app.post("/api/process")
async def process_reports(
    request: Request,
    background_tasks: BackgroundTasks,
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...)
):
    # Verify Authentication & Active Status
    current_user = get_current_user(request)
    print(f"Processing report request for authenticated user: {current_user['username']} ({current_user['role']})")

    if not pdf_file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Left file must be a PDF statement.")
        
    if not excel_file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Right file must be an Excel spreadsheet (.xlsx).")
        
    session_id = str(uuid.uuid4())
    temp_pdf_path = os.path.join(TEMP_DIR, f"{session_id}_statement.pdf")
    temp_excel_path = os.path.join(TEMP_DIR, f"{session_id}_template.xlsx")
    temp_output_path = os.path.join(TEMP_DIR, f"{session_id}_output.xlsx")
    
    try:
        with open(temp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(pdf_file.file, buffer)
            
        with open(temp_excel_path, "wb") as buffer:
            shutil.copyfileobj(excel_file.file, buffer)
            
        print(f"[{session_id}] Extracting PDF Data...")
        parsed_data = parse_pdf(temp_pdf_path, temp_excel_path)
        
        if not parsed_data["funds"]:
            raise ValueError("No Manulife fund investment records could be found or extracted from the PDF statement.")
            
        print(f"[{session_id}] Matching Fund Records and Calculating fields...")
        matched_count = update_excel_report(temp_excel_path, temp_output_path, parsed_data)
        
        if matched_count == 0:
            raise ValueError("No matching fund names found between the PDF statement and the Excel template. Please check the fund name column in your template.")
            
        background_tasks.add_task(cleanup_temp_files, temp_pdf_path, temp_excel_path, temp_output_path)
        
        import re
        account_holder = parsed_data.get("account_holder", "Customer").strip()
        clean_holder = re.sub(r'[\\/*?:"<>|]', "", account_holder)
        download_name = f"FNCL Group - {clean_holder}_Report.xlsx"
        
        return FileResponse(
            path=temp_output_path,
            filename=download_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        cleanup_temp_files(temp_pdf_path, temp_excel_path, temp_output_path)
        print(f"Error processing files: {e}")
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)}
        )

# Serve Frontend Static Files
frontend_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "frontend")
os.makedirs(frontend_dir, exist_ok=True)

app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="static")
