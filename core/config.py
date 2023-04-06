import os
from pathlib import Path
from dotenv import load_dotenv, find_dotenv

env_path = Path('.') / '.env'
load_dotenv(dotenv_path='C:/Users/jbailey/Documents/Python Projects/smartsheet_purchasing/.env')
# or load from the first .env file found
load_dotenv(dotenv_path=find_dotenv())

class Settings:
    """Class representing the settings used for interacting with Smartsheet API and Excel Sheets.

    Attributes:
        SMARTSHEET_API_URL (str): The base URL for accessing the Smartsheet API.
        API_KEY (str): The API key used for authenticating requests to the Smartsheet API.
        TEMPLATE_SHEET (str): The ID of the Smartsheet template sheet to use.
        WORKSPACE_ID (str): The ID of the Smartsheet workspace to use for production.
        TEST_API_KEY (str): The API key used for authenticating requests to the Smartsheet API during testing.
        TEST_TEMPLATE_ID (str): The ID of the Smartsheet template sheet to use during testing.
        TEST_WORKSPACE_ID (str): The ID of the Smartsheet workspace to use during testing.
        EXCEL_TAB (str): The name of the tab in the Excel sheet to use.
        TABLE_NAME (str): The name of the table in the Excel sheet to use.
    """
    # - SmartSheet Urls
    SMARTSHEET_API_URL = "https://api.smartsheet.com/2.0/sheets/"

    # - Production Credentials
    API_KEY = os.getenv("SMARTSHEET_API_KEY")
    TEMPLATE_SHEET = os.getenv("TEMPLATE_SHEET_ID")
    WORKSPACE_ID = os.getenv("LIVE_WORKSPACE_ID")

    # - Basic Testing Credentials
    TEST_API_KEY = os.getenv("TEST_API_KEY")
    TEST_TEMPLATE_ID = os.getenv("TEST_TEMPLATE_ID")
    TEST_WORKSPACE_ID = os.getenv("TEST_WORKSPACE_ID")

    # - Excel Sheet Info
    EXCEL_TAB = os.getenv("EXCEL_TAB")
    TABLE_NAME = os.getenv("TABLE_NAME")


settings = Settings()