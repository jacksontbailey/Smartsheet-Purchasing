from config import settings
from smartsheet_classes.smartsheet_api import SmartSheetApi


def create_new_smartsheet(sheet_name):
    try:

        new_sheet = SmartSheetApi(
            api_key = settings.API_KEY, 
            workspace_id = int(settings.WORKSPACE_ID)
        )
        
        print(f"Workspace is: {new_sheet.workspace_id}")
        
        new_sheet.create_sheet_in_workspace(
            new_sheet_name = sheet_name, 
            template_id = settings.TEMPLATE_SHEET
        )
    
    except Exception as e:
        print(str(e))