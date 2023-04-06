
# Smartsheet API Utility Module

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](#license)

This Python module provides utility functions and classes for working with the Smartsheet API.

## Description

The functions in this module can be used to create, read, write, and delete the sheets and the sheets' data in Smartsheets. It can also be used to read data from an Excel file and add it to an existing sheet. Additionally, it creates a GUI interface for the user using PySimpleGUI.

### Table of Contents

1. [Authors](#authors)
2. [Installation](#installation)
3. [Functions](#functions)
4. [Classes](#classes)
5. [Environment Variables](#environment-variables)
6. [License](#license)

## Authors

- [@JacksonTBailey](https://www.github.com/JacksonTBailey)

## Installation

The project has several dependencies, which are listed in the `requirements.txt` file. Some of these packages are not used directly, but are required as dependencies for other packages. It is recommended to use a virtual environment (venv) to install these dependencies.

1. Clone the repository to your local machine.

2. Install the required packages listed in `requirements.txt` using `pip install -r requirements.txt`. Note that not all of these packages are used directly, some of them come with other packages that were installed.

3. Create a virtual environment for your project using your preferred method.

4. Create a `.env` file in the root directory of the project and add all of the variables found in the [Environment Variables](#environment-variables) section.

5. Create a config.ini file in the root directory of the project and add the following variables:

```makefile
[GUI]
title = Your GUI's title
font_size = 10
titlebar_font_size = 11
font_family = Arial
theme = DarkTeal10

[COLORS]
dark_blue = #404563
lighter_blue = #51587B
white = #F3F3F7
french_gray = #B5B9CF

[EXCEL]
sheet_name = Purchasing_Items
```

This file stores your GUI's style settings. You can learn more about some of the various pre-made themes and additional styling through the [PySimpleGUI docs](https://www.pysimplegui.org/en/latest/)

## Functions

The module contains the following functions:

- **`create_new_smartsheet()`**: This function creates a new Smartsheet by calling the `create_sheet_in_workspace()` method of the `SmartsheetAPI` class. The function takes the Smartsheet client, the ID of the template sheet, and the name of the new sheet as arguments.
- **`create_new_smartsheet_window()`**: Display a window for creating a new Smartsheet and process the user input for creating a new sheet.
- **`create_smartsheet_client()`**: This function creates a Smartsheet client object by calling the `Smartsheet()` constructor of the Smartsheet SDK library.
- **`import_excel_data()`**: This function imports data from an Excel file by calling the `read_excel()` method of the pandas library.
- **`import_sheet_window(btn, import_function, option)`**: Display an import sheet window and process the user input for importing or updating data in Smartsheets.
- **`main()`**: Initializes and displays the main GUI window. Defines the layout for the GUI and creates the event loop that waits for user input.
- **`resource_path(relative_path)`**: Get absolute path to resource, works for dev and for PyInstaller.
- **`update_smartsheet()`**: This function updates data in Smartsheet by calling the `update_rows()` method of the `SmartsheetAPI` class.

## Classes

The module contains the following classes:

- **`ExcelSheetManager`**: A class that manages an Excel sheet. It allows users to check if the Excel file is open or not, check if the tab name exists in the Excel file, and read data from the specified tab in the Excel file.
- **`InputQuestions`**: A class that contains methods to prompt the user with questions to retrieve data and options. It prompts the user with a menu of options to create a new Smartsheet or import data into an existing one. The class also prompts the user to enter the name of the new Smartsheet, the name of the Excel sheet they want to import data from, and the name of the Smartsheet they want to import data into. The class also contains methods to start the timer to measure the runtime of the program, print the runtime of the program, and end the program if the user has indicated they want to end the program.
- **`Settings`**: A class representing the settings used for interacting with the Smartsheet API and Excel Sheets.
- **`SmartsheetApi`**: This class provides methods to interact with the Smartsheet API. You can create a SmartsheetAPI object by passing the API key, folder ID, sheet data, sheet ID, sheet name, and workspace ID as arguments. The class has various methods for retrieving, creating, updating, and deleting data from sheets. The methods include `get_sheets_in_folder()`, `get_sheets_in_workspace()`, `get_sheet_id_by_name()`, `create_sheet_in_workspace()`, `create_sheet_in_folder()`, `get_columns()`, `update_columns()`, `get_row()`, `add_rows()`, `update_rows()`, `delete_rows()`, `get_data()`, `check_duplicates()`, and `compare_data()`.

## Usage

To use this module, you will need to have a Smartsheet account and an API access token. You can get an API access token by following the instructions on the [Smartsheet API documentation](https://smartsheet-platform.github.io/api-docs/).

Here's a step-by-step guide on how to use this module:

### Step 1: Install the Required Libraries

Ensure that you have the required libraries installed, namely pandas, PySimpleGUI, and Smartsheet SDK. For ease of use, I have included a requirements.txt file. You can view the instructions for this in the [Installation](#installation) section of this document.

### Step 2: Set Up Your Environment

- Obtain an API token from Smartsheet and store it in a ".env" file in the project directory. Check out the [Environment Variables](#environment-variables) section for more details.

- Define the necessary settings for your Smartsheet API and Excel sheet interactions in the Settings class in the `config.py` file.

    ```python
        from config import settings

        # Load API keys and sheet information from .env file
        SMARTSHEET_API_KEY = settings.API_KEY
        TEMPLATE_SHEET_ID = settings.TEMPLATE_SHEET
        LIVE_WORKSPACE_ID = settings.WORKSPACE_ID
        TEST_API_KEY = settings.TEST_API_KEY
        TEST_TEMPLATE_ID = settings.TEST_TEMPLATE_ID
        TEST_WORKSPACE_ID = settings.TEST_WORKSPACE_ID
        EXCEL_TAB_NAME = settings.EXCEL_TAB
        TABLE_NAME = settings.TABLE_NAME
    ```

- Ensure that your Excel file is saved and closed before running the program to avoid any read/write errors.

### Step 3: Use the Main Function

Run the `main()` function in the run.py file. This will initialize and display the main GUI window. You can use this window to create a new Smartsheet, import data into an existing sheet, or exit the program.

### Step 4: Create a New Smartsheet

- Click the "Create a New Smartsheet" button on the main window.
  - A new window will appear asking you to input the name of the new Smartsheet.
- Enter the name and click the "Create" button.
  - The program will create a new Smartsheet using the specified name and the template sheet ID defined in the Settings class.

### Step 5: Import Data into an Existing Smartsheet

- Click the "Import Data into an Existing Smartsheet" button on the main window.
  - A new window will appear asking you to select the Excel file you want to import data from.
- Select the file and click the "Import" button.
  - A new window will appear asking you to select the Smartsheet you want to import data into.
- Select the sheet and click the "Import" button.
  - The program will import the data from the specified Excel sheet into the specified Smartsheet.

### Step 6: Exit the Program

- Click the "Exit" button on the main window.
  - The program will end.

That's it! With these steps, you can use this module to interact with the Smartsheet API and manage your Smartsheet data with ease.

## Environment Variables

To run this project, you will need to add the following environment variables to your .env file

- `SMARTSHEET_API_KEY`=your_api_key_here
- `TEMPLATE_SHEET_ID`=your_template_sheet_id_here
- `LIVE_WORKSPACE_ID`=your_live_workspace_id_here
- `TEST_API_KEY`=your_test_api_key_here
- `TEST_TEMPLATE_ID`=your_test_template_sheet_id_here
- `TEST_WORKSPACE_ID`=your_test_workspace_id_here
- `EXCEL_TAB`=your_excel_tab_name_here
- `TABLE_NAME`=your_table_name_here

## License

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
