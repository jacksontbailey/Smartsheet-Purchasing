import os
from pathlib import Path
from dotenv import load_dotenv

env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

class Settings:
    # - Production Credentials
    API_KEY = os.getenv("SMARTSHEET_API_KEY")
    TEMPLATE_SHEET = os.getenv("TEMPLATE_SHEET_ID")
    WORKSPACE_ID = os.getenv("WORKSPACE_ID")

    # - Basic Testing Credentials
    TEST_API_KEY = os.getenv("TEST_API_KEY")
    TEST_TEMPLATE_ID = os.getenv("TEST_TEMPLATE_ID")
    TEST_WORKSPACE_ID = os.getenv("TEST_WORKSPACE_ID")

    # - Fillable data at runtime
    NEW_SHEET_NAME = input("Enter sheet name: ")


settings = Settings()