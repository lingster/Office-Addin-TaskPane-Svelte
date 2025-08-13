import time
import json
from typing import Dict, Any
import aiohttp
from jose import jwt, jwk
from jose.exceptions import JOSEError
from fastapi import HTTPException, status

from .config import settings

class TokenCache:
    def __init__(self, ttl: int):
        self.ttl = ttl
        self._cache: Dict[str, Any] = {}
        self._expiry_time: float = 0

    def get(self) -> Any | None:
        if self._expiry_time > time.time():
            return self._cache
        return None

    def set(self, data: Any):
        self._cache = data
        self._expiry_time = time.time() + self.ttl

class TokenValidator:
    def __init__(self):
        self.jwks_url = f"{settings.AZURE_AD_AUTHORITY}/discovery/v2.0/keys"
        self.jwks_cache = TokenCache(ttl=settings.TOKEN_CACHE_TTL)
        self.algorithms = ["RS256"]

    async def get_signing_key(self, token: str) -> Dict[str, Any]:
        try:
            unverified_header = jwt.get_unverified_header(token)
        except JOSEError as e:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=f"Invalid token header: {e}")

        kid = unverified_header.get("kid")
        if not kid:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token header must contain 'kid'")

        jwks = self.jwks_cache.get()
        if not jwks:
            async with aiohttp.ClientSession() as session:
                try:
                    async with session.get(self.jwks_url) as response:
                        response.raise_for_status()
                        jwks_data = await response.json()
                        self.jwks_cache.set(jwks_data)
                except aiohttp.ClientError as e:
                    raise HTTPException(status_code=status.HTTP_503_SERVICE_UNAVAILABLE, detail=f"Could not fetch JWKS from Azure AD: {e}")

        for key in jwks["keys"]:
            if key["kid"] == kid:
                return key

        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=f"Signing key with kid '{kid}' not found in JWKS")

    async def validate_token(self, token: str) -> Dict[str, Any]:
        signing_key = await self.get_signing_key(token)

        try:
            public_key = jwk.construct(signing_key)
            payload = jwt.decode(
                token,
                public_key.to_pem(),
                algorithms=self.algorithms,
                audience=settings.API_AUDIENCE,
                issuer=f"https://sts.windows.net/{settings.TENANT_ID}/" # Note the trailing slash is important for some validations
            )
        except jwt.ExpiredSignatureError:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token has expired")
        except jwt.JWTClaimsError as e:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=f"Invalid claims: {e}")
        except JOSEError as e:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=f"Token validation failed: {e}")

        # Additional custom validation
        if 'nbf' in payload and time.time() < payload['nbf']:
             raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token not yet valid")

        if 'iat' in payload and (time.time() - payload['iat']) > settings.MAX_TOKEN_AGE:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token is too old")

        if payload.get("tid") != settings.TENANT_ID:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid tenant ID")

        return payload

token_validator = TokenValidator()
