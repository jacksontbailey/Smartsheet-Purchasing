import sys
import time

from config import settings
from smartsheet_classes.sheet_creator import SheetCreator


def main():
    try:
        # - Grab Currrent Time Before Running the Code
        start = time.time()

        template_sheet_id = settings.TEMPLATE_SHEET
        new_sheet_name = settings.NEW_SHEET_NAME
        destination_workspace_id = int(settings.WORKSPACE_ID)

        sheet_creator = SheetCreator(template_sheet_id, new_sheet_name, destination_workspace_id)
        new_sheet = sheet_creator.create_sheet()

        print(f"result: {new_sheet}")

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