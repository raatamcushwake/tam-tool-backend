from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    PROJECT_ID: str
    FIREBASE_CREDENTIAL_PATH: str
    SECRET_KEY: str
    ENVIRONMENT: str = "development"

    class Config:
        env_file = ".env"

settings = Settings()