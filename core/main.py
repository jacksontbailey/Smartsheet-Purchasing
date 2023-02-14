import sys

from input_questions import InputQuestions
from smartsheet_functions.import_excel_data import import_excel_data
from smartsheet_functions.create_new_sheet import create_new_smartsheet


def main():
    """
    The main function serves as the entry point for the script and provides a user-friendly interface to interact with the functionality provided by the code.
    """
        
    iq = InputQuestions()
    while not iq.end:
        try:
            # - Grab Currrent Time Before Running the Code
            option_selected = iq.main_script_options()

            if option_selected == 'Create New SmartSheet':
                NEW_SHEET_NAME = iq.name_new_smartsheet()
                create_new_smartsheet(sheet_name = NEW_SHEET_NAME)

            elif option_selected == 'Import Data into Existing Sheet':
                import_excel_data(import_question = iq)

            else:
                print("Incorrect input")


        except Exception as e:
            print(f"Error has occured: {e}")

    sys.exit(f"Finished Smartsheet Purchasing Sheet Creation.")


if __name__ == "__main__":
    main()