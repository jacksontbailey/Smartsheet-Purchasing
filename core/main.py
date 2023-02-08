import sys
import time

from config import settings
from smartsheet_classes.sheet_creator import SheetCreator
from smartsheet_classes.sheet_manager import SheetManager
from excel_classes.sheet_manager import ExcelSheetManager
from smartsheet_functions.insert_data_to_smartsheet import insert_data_to_smartsheet


def main():
    try:
        # - Grab Currrent Time Before Running the Code
        start = time.time()

        template_sheet_id = settings.TEMPLATE_SHEET
        new_sheet_name = settings.NEW_SHEET_NAME
        destination_workspace_id = int(settings.WORKSPACE_ID)
        api_key = settings.API_KEY
        excel_file = settings.TARGET_EXCEL_FILE
        folder = settings.EXCEL_FOLDER
        tab = settings.EXCEL_TAB
        excel_table = settings.TABLE_NAME

        sheet_creator = SheetCreator(api_key, template_sheet_id, new_sheet_name, destination_workspace_id)
        new_sheet = sheet_creator.create_sheet()
        print(f"result: {new_sheet}")

        excel_data = ExcelSheetManager(file_name=excel_file, file_location=folder, tab_name=tab, table_name=excel_table)
        
        insert_data_to_smartsheet(
            excel_data = excel_data.read_data(),
            sheet_manager = SheetManager(sheet_id=new_sheet.id, api_key=api_key),
        )


        # - Grab Currrent Time After Running the Code
        end = time.time()
        
        # - Subtract Start Time from The End Time
        total_time = end - start
        print(f"Total time to run script: {str(total_time)} seconds.")

        sys.exit(f"Finished Smartsheet Purchasing Sheet Creation.")


    except Exception as e:
        print(f"Error has occured: {e}")


if __name__ == "__main__":
    main()