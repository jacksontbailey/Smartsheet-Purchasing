import os
from pathlib import Path
from dotenv import load_dotenv

env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

class Settings:
    # - SmartSheet Urls
    SMARTSHEET_API_URL = "https://api.smartsheet.com/2.0/sheets/"

    # - Production Credentials
    API_KEY = os.getenv("SMARTSHEET_API_KEY")
    TEMPLATE_SHEET = os.getenv("TEMPLATE_SHEET_ID")
    WORKSPACE_ID = os.getenv("WORKSPACE_ID")

    # - Basic Testing Credentials
    TEST_API_KEY = os.getenv("TEST_API_KEY")
    TEST_TEMPLATE_ID = os.getenv("TEST_TEMPLATE_ID")
    TEST_WORKSPACE_ID = os.getenv("TEST_WORKSPACE_ID")

    # - Excel Sheet Info
    EXCEL_FOLDER = os.getenv("EXCEL_FOLDER")
    EXCEL_TAB = os.getenv("EXCEL_TAB")
    TABLE_NAME = os.getenv("TABLE_NAME")


settings = Settings()