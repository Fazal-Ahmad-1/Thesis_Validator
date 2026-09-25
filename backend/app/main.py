import logging
import os
import tempfile
import threading
import time
from collections import defaultdict, deque
from pathlib import Path
from typing import Optional

import jwt
from jwt import PyJWKClient
from supabase import create_client

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware

from validator.validator import validate_document

from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger("thesis_validator")


# ---------------------------------------------------------------------------
# Supabase configuration
# ---------------------------------------------------------------------------

SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_SECRET_KEY = os.getenv("SUPABASE_SECRET_KEY", "").strip()
SUPABASE_JWKS_URL = os.getenv("SUPABASE_JWKS_URL", "").strip()

if not SUPABASE_URL or not SUPABASE_SECRET_KEY or not SUPABASE_JWKS_URL:
    raise RuntimeError(
        "Missing required Supabase environment variables."
    )

# Server-side Supabase client.
# IMPORTANT: SUPABASE_SECRET_KEY is never exposed to the frontend.
supabase_admin = create_client(
    SUPABASE_URL,
    SUPABASE_SECRET_KEY
)

# Used to verify Supabase access-token signatures.
jwks_client = PyJWKClient(SUPABASE_JWKS_URL)


# ---------------------------------------------------------------------------
# Upload limits
# ---------------------------------------------------------------------------

MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MB
UPLOAD_CHUNK_SIZE = 1024 * 1024     # 1 MB


# ---------------------------------------------------------------------------
# Rate limiting
# ---------------------------------------------------------------------------

# Lightweight in-memory abuse protection for the public validation endpoint.
RATE_LIMIT_REQUESTS = 10
RATE_LIMIT_WINDOW_SECONDS = 60 * 60  # 1 hour

_rate_limit_lock = threading.Lock()
_rate_limit_requests = defaultdict(deque)


def _get_client_ip(request: Request) -> str:
    """Get the real client IP when running behind Cloudflare/Render."""

    cloudflare_ip = request.headers.get("cf-connecting-ip")

    if cloudflare_ip:
        return cloudflare_ip.strip()

    forwarded_for = request.headers.get("x-forwarded-for")

    if forwarded_for:
        return forwarded_for.split(",")[0].strip()

    if request.client:
        return request.client.host

    return "unknown"


def _check_rate_limit(request: Request) -> None:
    """Allow at most RATE_LIMIT_REQUESTS per client IP in the rolling window."""

    now = time.monotonic()
    client_ip = _get_client_ip(request)
    cutoff = now - RATE_LIMIT_WINDOW_SECONDS

    with _rate_limit_lock:
        timestamps = _rate_limit_requests[client_ip]

        while timestamps and timestamps[0] <= cutoff:
            timestamps.popleft()

        if len(timestamps) >= RATE_LIMIT_REQUESTS:
            raise HTTPException(
                status_code=429,
                detail="Too many validation requests. Please try again later."
            )

        timestamps.append(now)

        # Remove stale IP entries to prevent unbounded memory growth.
        stale_ips = [
            ip
            for ip, entries in _rate_limit_requests.items()
            if not entries or entries[-1] <= cutoff
        ]

        for ip in stale_ips:
            del _rate_limit_requests[ip]


# ---------------------------------------------------------------------------
# Supabase authentication
# ---------------------------------------------------------------------------

def _get_authenticated_user_id(request: Request) -> Optional[str]:
    """
    Return the authenticated Supabase user ID when a valid Bearer token
    is provided.

    No Authorization header means the request is treated as a guest.

    If an Authorization header is provided but the token is invalid or
    expired, the request is rejected with HTTP 401.
    """

    authorization = request.headers.get("authorization")

    # No token = guest request.
    if not authorization:
        return None

    if not authorization.lower().startswith("bearer "):
        raise HTTPException(
            status_code=401,
            detail="Invalid authorization header."
        )

    token = authorization[7:].strip()

    if not token:
        raise HTTPException(
            status_code=401,
            detail="Missing access token."
        )

    try:
        # Get the public signing key corresponding to the JWT.
        signing_key = jwks_client.get_signing_key_from_jwt(token)

        # Verify the JWT signature and important claims.
        payload = jwt.decode(
            token,
            signing_key.key,
            algorithms=[signing_key.algorithm_name],
            audience="authenticated",
            issuer=f"{SUPABASE_URL}/auth/v1",
        )

        user_id = payload.get("sub")

        if not user_id:
            raise HTTPException(
                status_code=401,
                detail="Invalid authentication token."
            )

        return user_id

    except HTTPException:
        raise

    except Exception:
        logger.warning("Invalid Supabase access token.")

        raise HTTPException(
            status_code=401,
            detail="Invalid or expired authentication token."
        )


# ---------------------------------------------------------------------------
# CORS
# ---------------------------------------------------------------------------

LOCAL_ORIGINS = [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
]

frontend_url = os.getenv("FRONTEND_URL", "").strip().rstrip("/")

ALLOWED_ORIGINS = LOCAL_ORIGINS.copy()

if frontend_url:
    ALLOWED_ORIGINS.append(frontend_url)


# ---------------------------------------------------------------------------
# Security headers
# ---------------------------------------------------------------------------

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    """Add basic browser security headers to API responses."""

    async def dispatch(self, request, call_next):
        response = await call_next(request)

        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"

        return response


# ---------------------------------------------------------------------------
# FastAPI application
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Basic endpoints
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# DOCX validation endpoint
# ---------------------------------------------------------------------------

@app.post("/api/validate")
async def validate_docx(
    request: Request,
    file: Optional[UploadFile] = File(default=None),
    profile_name: str = Form(...)
):
    """
    Validate an uploaded DOCX file against a named university rule profile.

    Guest users:
        - Can validate normally.
        - Validation result is returned but not saved.

    Authenticated users:
        - Can validate normally.
        - Validation metadata is saved to Supabase.

    Uploaded documents are written only to a temporary server-side file,
    processed, and removed in the finally block.
    """

    # -----------------------------------------------------------------------
    # 1. Rate limiting
    # -----------------------------------------------------------------------

    _check_rate_limit(request)

    # -----------------------------------------------------------------------
    # 2. Identify authenticated user
    # -----------------------------------------------------------------------

    user_id = _get_authenticated_user_id(request)

    # -----------------------------------------------------------------------
    # 3. Validate upload
    # -----------------------------------------------------------------------

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

        # -------------------------------------------------------------------
        # 4. Save upload temporarily
        # -------------------------------------------------------------------

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

        # -------------------------------------------------------------------
        # 5. Run the existing document validator
        # -------------------------------------------------------------------

        result = validate_document(
            str(tmp_path),
            profile_name
        )

        # -------------------------------------------------------------------
        # 6. Save validation metadata for authenticated users
        # -------------------------------------------------------------------

        if user_id:
            try:
                supabase_admin.table("validations").insert({
                    "user_id": user_id,
                    "profile_name": profile_name,
                    "score": result.get("score"),
                    "blocking_error_count": result.get(
                        "blocking_error_count",
                        0
                    ),
                }).execute()

            except Exception:
                # Do not make a successful document validation fail merely
                # because the optional history save failed.
                logger.exception(
                    "Could not save validation history for user %s",
                    user_id
                )

        # -------------------------------------------------------------------
        # 7. Return the original validation response
        # -------------------------------------------------------------------

        return result

    # -----------------------------------------------------------------------
    # Error handling
    # -----------------------------------------------------------------------

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

        raise HTTPException(
            status_code=400,
            detail=str(exc)
        )

    except Exception:
        logger.exception(
            "Unexpected error while validating upload '%s'",
            file.filename
        )

        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while validating the document."
        )

    # -----------------------------------------------------------------------
    # Always remove temporary document
    # -----------------------------------------------------------------------

    finally:

        if tmp_path is not None and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                logger.exception(
                    "Could not remove temporary upload file."
                )

        await file.close()