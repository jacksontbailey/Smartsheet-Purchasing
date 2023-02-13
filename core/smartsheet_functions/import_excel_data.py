from config import settings
from excel_classes.sheet_manager import ExcelSheetManager
from smartsheet_classes.smartsheet_api import SmartSheetApi
from smartsheet_functions.insert_data_to_smartsheet import insert_data_to_smartsheet



def import_excel_data(import_question):
    TARGET_EXCEL_FILE = import_question.select_excel_sheet()

    excel_data = ExcelSheetManager(
        file_name = TARGET_EXCEL_FILE, 
        file_location = settings.EXCEL_FOLDER, 
        tab_name = settings.EXCEL_TAB, 
        table_name = settings.TABLE_NAME
    )

    TARGET_SHEET = import_question.select_smartsheet_name()

    sheet_manager = SmartSheetApi(
        api_key = settings.API_KEY, 
        sheet_name = TARGET_SHEET,
        workspace_id = int(settings.WORKSPACE_ID)
    )

    insert_data_to_smartsheet(
            excel_data = excel_data.read_data(),
            sheet_manager = SmartSheetApi(
                api_key = settings.API_KEY, 
                sheet_id = sheet_manager.get_sheet_id_by_name()
            ),
        )
