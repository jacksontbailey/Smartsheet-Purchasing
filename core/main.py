"""This module provides utility functions and classes for working with the SmartSheet API.

The functions in this module can be used to create, read, write, and delete
the sheets and the sheets' data in SmartSheets. It can also be used to read data 
from an Excel file and add it to an existing sheet.


The module contains the following functions:
    - create_new_smartsheet
    - create_smartsheet_client
    - import_excel_data
    - insert_data_to_smartsheet

    
The module contains the following classes:
    - ExcelSheetManager
    - InputQuestions
    - Settings
    - SmartSheetApi

Author:
    @ Jackson Bailey - 2023.02.14

License:
    MIT License

    Copyright (c) 2023 Jackson Bailey

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

"""

import sys

from input_questions import InputQuestions
from smartsheet_functions.create_new_sheet import create_new_smartsheet
from smartsheet_functions.import_excel_data import import_excel_data


def main():
    """
    The main function serves as the entry point for the script and provides \
    a user-friendly interface to interact with the functionality provided by the code.
    """
    input_question = InputQuestions()
    while not input_question.end:
        try:
            # - Grab Currrent Time Before Running the Code
            option_selected = input_question.main_script_options()

            if option_selected == 'Create New SmartSheet':
                new_sheet_name = input_question.name_new_smartsheet()
                create_new_smartsheet(sheet_name = new_sheet_name)

            elif option_selected == 'Import Data into Existing Sheet' or 'Update Existing SmartSheet with Excel File':
                import_excel_data(input_question = input_question, option = option_selected)
            
            else:
                print("Incorrect input")


        except FileNotFoundError as caught_error:
            print(f"File not found: {caught_error}")
        except ValueError as caught_error:
            print(f"Value error: {caught_error}")
        except Exception as caught_error:
            print(f"Unexpected error: {caught_error}")

    sys.exit("Finished Smartsheet Purchasing Sheet Creation.")


if __name__ == "__main__":
    main()
