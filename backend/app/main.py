import logging
import os
import tempfile
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware

from validator.validator import validate_document

logger = logging.getLogger("thesis_validator")

MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MB
UPLOAD_CHUNK_SIZE = 1024 * 1024     # 1 MB

LOCAL_ORIGINS = [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
]

frontend_url = os.getenv("FRONTEND_URL", "").strip().rstrip("/")
ALLOWED_ORIGINS = LOCAL_ORIGINS.copy()

if frontend_url:
    ALLOWED_ORIGINS.append(frontend_url)


class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    """Add basic browser security headers to API responses."""

    async def dispatch(self, request, call_next):
        response = await call_next(request)

        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"

        return response


app = FastAPI(
    title="Thesis Validator API",
    description="API for checking thesis and report formatting",
    version="1.0.0"
)

app.add_middleware(SecurityHeadersMiddleware)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


@app.get("/")
def root():
    return {
        "message": "Thesis Validator API is running"
    }


@app.get("/api/health")
def health_check():
    return {
        "status": "healthy"
    }


@app.post("/api/validate")
async def validate_docx(
    file: Optional[UploadFile] = File(default=None),
    profile_name: str = Form(...)
):
    """
    Validate an uploaded DOCX file against a named university rule profile.

    Uploaded documents are written only to a temporary server-side file,
    processed, and removed in the finally block.
    """

    if file is None or not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file was uploaded. Please attach a .docx file."
        )

    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Only .docx files are supported."
        )

    tmp_path = None
    total_size = 0

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".docx",
            delete=False
        ) as tmp_file:
            tmp_path = Path(tmp_file.name)

            while True:
                chunk = await file.read(UPLOAD_CHUNK_SIZE)

                if not chunk:
                    break

                total_size += len(chunk)

                if total_size > MAX_UPLOAD_SIZE:
                    raise HTTPException(
                        status_code=413,
                        detail="Uploaded file is too large. Maximum allowed size is 10 MB."
                    )

                tmp_file.write(chunk)

        result = validate_document(str(tmp_path), profile_name)
        return result

    except HTTPException:
        raise

    except FileNotFoundError:
        raise HTTPException(
            status_code=400,
            detail=f"Unknown rule profile: '{profile_name}'."
        )

    except ValueError as exc:
        logger.warning(
            "Validation input error for '%s' (profile=%s): %s",
            file.filename,
            profile_name,
            exc
        )
        raise HTTPException(status_code=400, detail=str(exc))

    except Exception:
        logger.exception(
            "Unexpected error while validating upload '%s'",
            file.filename
        )
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while validating the document."
        )

    finally:
        if tmp_path is not None and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                logger.exception("Could not remove temporary upload file.")

        await file.close()
