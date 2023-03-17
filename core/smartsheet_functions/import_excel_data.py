from core.config import settings
from core.excel_classes.sheet_manager import ExcelSheetManager
from core.smartsheet_classes.smartsheet_api import SmartSheetApi
from core.smartsheet_functions.insert_data_to_smartsheet import insert_data_to_smartsheet
from core.smartsheet_functions.update_data_in_smartsheet import update_smartsheet


def import_excel_data(option, url, smartsheet_name):
    """
    Imports data from an Excel sheet into a Smartsheet.
    
    This function imports data from an Excel sheet into a Smartsheet. It does so by first selecting the Excel sheet from which to grab data from and the Smartsheet to which to upload that data. It then reads in the data from the selected Excel sheet and inserts it into the selected Smartsheet.
    
    Parameters:
        input_question (object): an instance of the `InputQuestions` class that contains information about the selected options for the import process.
    
    Returns:
        None: This function does not return any value.
    
    """

    excel_data = ExcelSheetManager(
        filebrowse_url = url,
        tab_name = settings.EXCEL_TAB, 
        table_name = settings.TABLE_NAME
    )


    sheet_manager = SmartSheetApi(
        api_key = settings.API_KEY, 
        sheet_name = smartsheet_name,
        workspace_id = int(settings.WORKSPACE_ID)
    )

    sheet_manager.get_sheet_id_by_name()

    if option == 'import_sheet':
        insert_data_to_smartsheet(
                excel_data = excel_data.read_data(),
                sheet_manager = sheet_manager
            )
    
    else:
        update_smartsheet(
            excel_data = excel_data.read_data(),
            sheet_manager = sheet_manager
        )
