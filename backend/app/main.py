import logging
import json
from datetime import datetime
from fastapi import FastAPI, HTTPException, Response, status
from fastapi.middleware.cors import CORSMiddleware

from .config import settings
from .auth import token_validator
from .models import TokenRequest, UserInfo, AuthError

# Configure logging
logging.basicConfig(level=settings.LOG_LEVEL)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Office Add-in SSO Backend",
    description="Handles SSO authentication for the Office Add-in.",
    version="2.0.0",
    debug=settings.DEBUG
)

# Configure CORS
origins = [origin.strip() for origin in settings.CORS_ORIGINS.split(',')]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def log_auth_event(event_type: str, user_oid: str = None, details: dict = None):
    """Log authentication events with a consistent structure."""
    log_data = {
        "timestamp": datetime.utcnow().isoformat(),
        "event_type": event_type,
        "user_oid": user_oid,
        "details": details or {}
    }
    logger.info(json.dumps(log_data))


@app.post("/api/auth/microsoft", response_model=UserInfo, responses={401: {"model": AuthError}})
async def authenticate_microsoft(token_request: TokenRequest):
    """
    Validates an Azure AD SSO token and returns user information.
    """
    try:
        payload = await token_validator.validate_token(token_request.token)

        user_info = UserInfo(
            user=payload.get("name"),
            email=payload.get("preferred_username"),
            oid=payload.get("oid"),
            tenant=payload.get("tid")
        )

        log_auth_event("authentication_success", user_oid=user_info.oid)
        return user_info

    except HTTPException as e:
        log_auth_event("authentication_failure", details={"status_code": e.status_code, "detail": e.detail})
        # Re-raise the exception to let FastAPI handle the response
        raise e
    except Exception as e:
        log_auth_event("authentication_error", details={"error": str(e)})
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An unexpected error occurred during authentication."
        )

@app.get("/api/health")
async def health_check():
    """
    Simple health check endpoint.
    """
    # In a real app, this would check DB connection, etc.
    # For now, we'll just check Azure AD connectivity as a basic measure.
    is_healthy = True
    azure_ad_status = "healthy"
    try:
        # A simple check, doesn't validate content, just connectivity
        await token_validator.get_signing_key("dummy")
    except HTTPException as e:
        # We expect a 401 for a dummy token, that's ok.
        # But a 503 means Azure AD is unreachable.
        if e.status_code == status.HTTP_503_SERVICE_UNAVAILABLE:
            is_healthy = False
            azure_ad_status = "unhealthy"

    health_status = {
        "status": "healthy" if is_healthy else "unhealthy",
        "timestamp": datetime.utcnow().isoformat(),
        "checks": {
            "azure_ad_connectivity": azure_ad_status
        }
    }

    status_code = 200 if is_healthy else 503
    return Response(content=json.dumps(health_status), status_code=status_code, media_type="application/json")
