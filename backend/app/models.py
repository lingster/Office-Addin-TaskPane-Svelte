from pydantic import BaseModel, Field
from datetime import datetime

class TokenRequest(BaseModel):
    token: str

class UserInfo(BaseModel):
    user: str | None = None
    email: str | None = None
    oid: str
    tenant: str
    authenticated_at: datetime = Field(default_factory=datetime.utcnow)

class AuthError(BaseModel):
    error: str
    error_description: str
    timestamp: datetime = Field(default_factory=datetime.utcnow)
