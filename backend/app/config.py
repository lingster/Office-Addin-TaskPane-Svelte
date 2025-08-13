import os
from pydantic_settings import BaseSettings
from dotenv import load_dotenv

# Correctly locate the .env.backend file relative to this script's location
# This assumes config.py is in backend/app/
dotenv_path = os.path.join(os.path.dirname(__file__), '..', '.env.backend')
load_dotenv(dotenv_path=dotenv_path)

class Settings(BaseSettings):
    TENANT_ID: str
    CLIENT_ID: str
    API_AUDIENCE: str
    AZURE_AD_AUTHORITY: str
    DEBUG: bool = True
    LOG_LEVEL: str = "INFO"
    CORS_ORIGINS: str
    TOKEN_CACHE_TTL: int = 3600
    MAX_TOKEN_AGE: int = 3600

    class Config:
        # Pydantic-settings will automatically look for a .env file,
        # but we specify it here for clarity and to ensure it's loaded.
        env_file = ".env.backend"
        env_file_encoding = "utf-8"
        # This is important for when running with Gunicorn or Uvicorn from the project root
        # and you want to load a specific env file.
        # However, load_dotenv above should handle it robustly.
        # We keep it for best practice.

# Instantiate the settings
settings = Settings()
