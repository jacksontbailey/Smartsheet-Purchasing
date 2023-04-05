from ..config import settings
from ..excel_classes.sheet_manager import ExcelSheetManager
from ..smartsheet_classes.smartsheet_api import SmartSheetApi
from ..smartsheet_functions.insert_data_to_smartsheet import insert_data_to_smartsheet
from ..smartsheet_functions.update_data_in_smartsheet import update_smartsheet


def import_excel_data(option, url, smartsheet_name):
    """
    Imports data from an Excel sheet into a Smartsheet.

    This function imports data from the specified Excel sheet into the specified Smartsheet. It first initializes an instance
    of the `ExcelSheetManager` class with the provided URL, tab name, and table name. It also initializes an instance of the 
    `SmartSheetApi` class with the provided API key, sheet name, and workspace ID. It then reads the data from the selected Excel
    sheet and inserts it into the selected Smartsheet using the `update_smartsheet` function. If the selected Excel sheet has an
    incorrect tab name or if the Smartsheet has an invalid ID, the function returns an error message.

    Parameters:
        option (object): an instance of the `InputQuestions` class that contains information about the selected options for 
        the import process.
        url (str): a string containing the URL of the Excel sheet to import data from.
        smartsheet_name (str): a string containing the name of the Smartsheet to import data into.

    Returns:
        None: This function does not return any value.
    """

    workspace = int(settings.WORKSPACE_ID)

    excel_data = ExcelSheetManager(
        filebrowse_url = url,
        tab_name = settings.EXCEL_TAB, 
        table_name = settings.TABLE_NAME
    )

    sheet_manager = SmartSheetApi(
        api_key = settings.API_KEY, 
        sheet_name = smartsheet_name,
        workspace_id = workspace
    )

    
    excel_data_validation = excel_data.read_data()

    if excel_data_validation == "Incorrect Tab Name":
        return [excel_data_validation]

    sheet_id = sheet_manager.get_sheet_id_by_name()

    if not sheet_id:
        return ["Invalid ID"]

    response = update_smartsheet(
        excel_data = excel_data.read_data(),
        sheet_manager = sheet_manager,
        option = option
    )
    
    return response
