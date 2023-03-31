from ..config import settings
from ..smartsheet_classes.smartsheet_api import SmartSheetApi


def create_new_smartsheet(sheet_name):
    """Create a new smartsheet.

    Args:
        sheet_name (str): The name of the new smartsheet to be created.

    Returns:
        None: This function does not return any value.
    """

    try:

        new_sheet = SmartSheetApi(
            api_key = settings.API_KEY, 
            workspace_id = int(settings.WORKSPACE_ID)
        )
        
        print(f"Workspace is: {new_sheet.workspace_id}")
        
        new_sheet.create_sheet_in_workspace(
            template_sheet_id = settings.TEMPLATE_SHEET,
            new_sheet_name = sheet_name, 
        )
    
    except Exception as e:
        print(str(e))