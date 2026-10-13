import os
import io
from typing import Optional, Any
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Query, Request
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse, RedirectResponse
from fastapi.encoders import jsonable_encoder
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.sessions import SessionMiddleware
import hmac
import zipfile
import numpy as np
import pandas as pd

from pydantic import BaseModel
from backend.excel_reader import load_all_excel_files
from backend.validator import validate_inputs, AllocationValidationError
from backend.scoring import calculate_compatibility_scores
from backend.agent import FacultyAllocationAgent
from input.sample_generator import generate_sample_files

app = FastAPI(title="AI Faculty-Subject-Section Allocation Agent API", version="1.0.0")

class ChatRequest(BaseModel):
    question: str

# Session-based login. Change these environment variables in production.
SESSION_SECRET = os.getenv("SESSION_SECRET", "change-this-session-secret-in-production")
LOGIN_USERNAME = os.getenv("FACULTY_APP_USERNAME", "admin")
LOGIN_PASSWORD = os.getenv("FACULTY_APP_PASSWORD", "admin123")
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET, max_age=60 * 60 * 8, same_site="lax", https_only=False)

def is_authenticated(request: Request) -> bool:
    return request.session.get("authenticated") is True

def verify_credentials(username: str, password: str) -> bool:
    return hmac.compare_digest(username, LOGIN_USERNAME) and hmac.compare_digest(password, LOGIN_PASSWORD)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Store persistent upload paths
UPLOAD_DIR = "input"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Ensure sample files exist
generate_sample_files(UPLOAD_DIR)

agent_orchestrator = FacultyAllocationAgent()

def sanitize_for_json(data: Any) -> Any:
    """Recursively converts numpy and pandas types to standard Python primitives for JSON response."""
    if isinstance(data, dict):
        return {str(k): sanitize_for_json(v) for k, v in data.items()}
    elif isinstance(data, (list, tuple, set)):
        return [sanitize_for_json(v) for v in data]
    elif isinstance(data, (np.integer, np.int64, np.int32)):
        return int(data)
    elif isinstance(data, (np.floating, np.float64, np.float32)):
        return float(data)
    elif isinstance(data, np.ndarray):
        return [sanitize_for_json(v) for v in data.tolist()]
    elif pd.isna(data):
        return None
    return data

@app.post("/api/login")
async def login(request: Request, username: str = Form(...), password: str = Form(...)):
    if not verify_credentials(username.strip(), password):
        accept = request.headers.get("accept", "")
        if "application/json" in accept or request.headers.get("x-requested-with") == "XMLHttpRequest":
            return JSONResponse(status_code=401, content={"success": False, "message": "Invalid username or password."})
        return RedirectResponse(url="/login?error=invalid", status_code=303)
    request.session.clear()
    request.session["authenticated"] = True
    request.session["username"] = username.strip()
    accept = request.headers.get("accept", "")
    if "application/json" in accept or request.headers.get("x-requested-with") == "XMLHttpRequest":
        return JSONResponse(content={"success": True, "message": "Login successful."})
    return RedirectResponse(url="/", status_code=303)

@app.get("/api/auth/me")
async def auth_me(request: Request):
    if not is_authenticated(request):
        return JSONResponse(status_code=401, content={"authenticated": False})
    return {"authenticated": True, "username": request.session.get("username", "User")}

@app.post("/api/logout")
async def logout(request: Request):
    request.session.clear()
    return {"success": True}

@app.post("/api/sample-data")
async def get_sample_files(request: Request):
    """Generates sample Excel files and returns them as a zip package."""
    if not is_authenticated(request):
        raise HTTPException(status_code=401, detail="Authentication required.")
    pref_path = os.path.join(UPLOAD_DIR, "faculty_preferences.xlsx")
    exp_path = os.path.join(UPLOAD_DIR, "faculty_experience.xlsx")
    course_path = os.path.join(UPLOAD_DIR, "course_details.xlsx")
    comb_path = os.path.join(UPLOAD_DIR, "faculty_sample_combined.xlsx")
    
    if not os.path.exists(comb_path) or not os.path.exists(course_path):
        generate_sample_files(UPLOAD_DIR)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        if os.path.exists(comb_path):
            zip_file.write(comb_path, arcname="faculty_sample_combined.xlsx")
        if os.path.exists(pref_path):
            zip_file.write(pref_path, arcname="faculty_preferences.xlsx")
        if os.path.exists(exp_path):
            zip_file.write(exp_path, arcname="faculty_experience.xlsx")
        if os.path.exists(course_path):
            zip_file.write(course_path, arcname="course_details.xlsx")
    
    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=faculty_sample_excel_files.zip"}
    )

@app.post("/api/validate")
async def validate_allocation_request(
    request: Request,
    num_sections: int = Form(...),
    num_subjects: int = Form(...),
    sections_per_faculty: int = Form(...),
    max_subjects_per_faculty: int = Form(...),
    dean_max_sections: int = Form(2),
    hod_max_sections: int = Form(3),
    high_exp_threshold: float = Form(8.0),
    faculty_file: Optional[UploadFile] = File(None),
    pref_file: Optional[UploadFile] = File(None),
    exp_file: Optional[UploadFile] = File(None),
    course_file: Optional[UploadFile] = File(None)
):
    """
    Validates uploaded Excel files and configuration inputs.
    Accepts single combined faculty file or legacy split files.
    """
    if not is_authenticated(request):
        raise HTTPException(status_code=401, detail="Authentication required.")
    try:
        target_fac = faculty_file or pref_file
        if target_fac:
            pref_src = target_fac.file
            exp_src = exp_file.file if exp_file else None
        elif exp_file:
            comb_sample = os.path.join(UPLOAD_DIR, "faculty_sample_combined.xlsx")
            pref_src = comb_sample if os.path.exists(comb_sample) else os.path.join(UPLOAD_DIR, "faculty_preferences.xlsx")
            exp_src = exp_file.file
        else:
            comb_sample = os.path.join(UPLOAD_DIR, "faculty_sample_combined.xlsx")
            if os.path.exists(comb_sample):
                pref_src = comb_sample
                exp_src = None
            else:
                pref_src = os.path.join(UPLOAD_DIR, "faculty_preferences.xlsx")
                exp_src = os.path.join(UPLOAD_DIR, "faculty_experience.xlsx")

        course_src = course_file.file if course_file else (
            os.path.join(UPLOAD_DIR, "course_details.xlsx") if os.path.exists(os.path.join(UPLOAD_DIR, "course_details.xlsx")) else None
        )

        pref_df, exp_df, course_df = load_all_excel_files(pref_src, exp_src, course_src)
        
        val_summary = validate_inputs(
            pref_df=pref_df,
            exp_df=exp_df,
            num_sections=num_sections,
            num_subjects=num_subjects,
            sections_per_faculty=sections_per_faculty,
            max_subjects_per_faculty=max_subjects_per_faculty,
            course_df=course_df,
            dean_max_sections=dean_max_sections,
            hod_max_sections=hod_max_sections,
            high_exp_threshold=high_exp_threshold
        )

        return JSONResponse(content=jsonable_encoder(sanitize_for_json({
            "success": True,
            "message": "Data & constraint validation passed successfully!",
            "summary": val_summary
        })))
    except AllocationValidationError as e:
        return JSONResponse(status_code=400, content=jsonable_encoder(sanitize_for_json({
            "success": False,
            "error_type": "ValidationError",
            "message": e.message,
            "details": e.details
        })))
    except Exception as e:
        return JSONResponse(status_code=500, content=jsonable_encoder(sanitize_for_json({
            "success": False,
            "error_type": "SystemError",
            "message": str(e)
        })))

@app.post("/api/allocate")
async def run_allocation(
    request: Request,
    num_sections: int = Form(...),
    num_subjects: int = Form(...),
    sections_per_faculty: int = Form(...),
    max_subjects_per_faculty: int = Form(...),
    dean_max_sections: int = Form(2),
    hod_max_sections: int = Form(3),
    high_exp_threshold: float = Form(8.0),
    weight_pref: float = Form(0.40),
    weight_exp: float = Form(0.30),
    weight_skill: float = Form(0.20),
    weight_qual: float = Form(0.10),
    faculty_file: Optional[UploadFile] = File(None),
    pref_file: Optional[UploadFile] = File(None),
    exp_file: Optional[UploadFile] = File(None),
    course_file: Optional[UploadFile] = File(None)
):
    """
    Runs full allocation pipeline: Validation -> Scoring -> OR-Tools Optimization -> Excel Export -> AI Explanation.
    Accepts single combined faculty file or legacy split files.
    """
    if not is_authenticated(request):
        raise HTTPException(status_code=401, detail="Authentication required.")
    try:
        target_fac = faculty_file or pref_file
        if target_fac:
            pref_src = target_fac.file
            exp_src = exp_file.file if exp_file else None
        elif exp_file:
            comb_sample = os.path.join(UPLOAD_DIR, "faculty_sample_combined.xlsx")
            pref_src = comb_sample if os.path.exists(comb_sample) else os.path.join(UPLOAD_DIR, "faculty_preferences.xlsx")
            exp_src = exp_file.file
        else:
            comb_sample = os.path.join(UPLOAD_DIR, "faculty_sample_combined.xlsx")
            if os.path.exists(comb_sample):
                pref_src = comb_sample
                exp_src = None
            else:
                pref_src = os.path.join(UPLOAD_DIR, "faculty_preferences.xlsx")
                exp_src = os.path.join(UPLOAD_DIR, "faculty_experience.xlsx")

        course_src = course_file.file if course_file else (
            os.path.join(UPLOAD_DIR, "course_details.xlsx") if os.path.exists(os.path.join(UPLOAD_DIR, "course_details.xlsx")) else None
        )

        custom_weights = {
            "preference": weight_pref,
            "experience": weight_exp,
            "skill": weight_skill,
            "qualification": weight_qual
        }

        output_path = os.path.join(OUTPUT_DIR, "faculty_allocation.xlsx")

        result = agent_orchestrator.run_pipeline(
            pref_file_source=pref_src,
            exp_file_source=exp_src,
            num_sections=num_sections,
            num_subjects=num_subjects,
            sections_per_faculty=sections_per_faculty,
            max_subjects_per_faculty=max_subjects_per_faculty,
            course_file_source=course_src,
            dean_max_sections=dean_max_sections,
            hod_max_sections=hod_max_sections,
            high_exp_threshold=high_exp_threshold,
            weights=custom_weights,
            output_excel_path=output_path
        )

        return JSONResponse(content=jsonable_encoder(sanitize_for_json(result)))

    except AllocationValidationError as e:
        return JSONResponse(status_code=400, content=jsonable_encoder(sanitize_for_json({
            "success": False,
            "error_type": "ValidationError",
            "message": e.message,
            "details": e.details
        })))
    except Exception as e:
        return JSONResponse(status_code=500, content=jsonable_encoder(sanitize_for_json({
            "success": False,
            "error_type": "AllocationError",
            "message": str(e)
        })))


@app.get("/api/download/faculty_allocation.xlsx")
async def download_allocation_excel(request: Request):
    """Downloads the generated faculty_allocation.xlsx file."""
    if not is_authenticated(request):
        raise HTTPException(status_code=401, detail="Authentication required.")
    file_path = os.path.join(OUTPUT_DIR, "faculty_allocation.xlsx")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Allocation Excel file has not been generated yet.")
    return FileResponse(
        path=file_path,
        filename="faculty_allocation.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=faculty_allocation.xlsx",
            "Cache-Control": "no-cache"
        }
    )

@app.post("/api/chat")
async def chat_with_agent(req: ChatRequest, request: Request):
    """
    Handles live natural language queries from Deans and HODs regarding allocation results,
    workload totals, assignment rationales, and scenario A/B comparison.
    """
    if not is_authenticated(request):
        raise HTTPException(status_code=401, detail="Authentication required.")
    try:
        chat_res = agent_orchestrator.handle_chat_query(req.question)
        return JSONResponse(content=jsonable_encoder(sanitize_for_json(chat_res)))
    except Exception as e:
        return JSONResponse(status_code=500, content=jsonable_encoder(sanitize_for_json({
            "success": False,
            "error_type": "ChatError",
            "message": str(e)
        })))

# Serve Frontend static assets
frontend_dir = os.path.abspath("frontend")
if os.path.exists(frontend_dir):
    app.mount("/static", StaticFiles(directory=frontend_dir), name="static")

@app.get("/")
async def serve_index(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=303)
    index_file = os.path.join(frontend_dir, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {"message": "Frontend index.html not found"}

@app.get("/login")
async def serve_login(request: Request):
    if is_authenticated(request):
        return RedirectResponse(url="/")
    return FileResponse(os.path.join(frontend_dir, "login.html"))
