"""This module provides utility functions and classes for working with the Smartsheet API.

The functions in this module can be used to create, read, write, and delete
the sheets and the sheets' data in Smartsheets. It can also be used to read data 
from an Excel file and add it to an existing sheet.


The module contains the following functions:
    - create_new_smartsheet
    - create_new_smartsheet_window
    - create_smartsheet_client
    - import_excel_data
    - import_sheet_window
    - main
    - update_smartsheet
    - resource_path
    
The module contains the following classes:
    - ExcelSheetManager
    - InputQuestions
    - Settings
    - SmartsheetApi

Author:
    @ Jackson Bailey - 2023.02.14

"""



from core.gui.create_smartsheet_window import create_new_smartsheet_window
from core.gui.import_sheet_window import import_sheet_window
from core.smartsheet_functions.import_excel_data import import_excel_data
from pathlib import Path
import PySimpleGUI as sg
import traceback
import sys
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def main():
    """
    Initializes and displays the main GUI window. Defines the layout for the GUI and creates the event loop that waits for user input.
    
    Parameters:
        None
    
    Returns:
        None
    """
    # define the layout for the main GUI window
    frame_layout = [
        [sg.Button("Import Data", border_width=0, font=('any 10 bold' ), k='-IMPORT-', tooltip='Import excel data into an empty Smartsheet', expand_x=True, p=((30, 7.5), (20, 10))),
         sg.Button("Update Sheet", border_width=0, font=('any 10 bold' ), k='-UPDATE-', tooltip='Update the existing Smartsheet with an excel file', expand_x=True, p=((7.5, 30), (20, 10)))],
        [sg.Button('Create Sheet', border_width=0, font=('any 10 bold' ), k='-NEW-', tooltip='Create a new Smartsheet', expand_x=True, p=(30,(10, 20)))],
    ]

    layout = [
        [sg.Frame('Smartsheet Actions', frame_layout, font='Any 16', expand_x=True, pad=((10), 0))],
        [sg.Push(), sg.Cancel(button_text="Exit", font=('any 10 bold' ), k='exit', s=10, pad=((0, 10), (10, 0)))]
    ]

    # create the main GUI window
    window_title = settings['GUI']['title']
    window = sg.Window(window_title, layout, finalize=True, size=(400, 225), return_keyboard_events=True)
    
    # set the cursor for each button
    window['-NEW-'].set_cursor(cursor="hand2")
    window['-UPDATE-'].set_cursor(cursor="hand2")
    window['-IMPORT-'].set_cursor(cursor="hand2")
    window['exit'].set_cursor(cursor="hand2")

    try:
        # event loop for the main GUI window
        while True:
            event, values = window.read()
            
            # exit the program if the window is closed or the user clicks the 'exit' button
            if event == sg.WINDOW_CLOSED or event == 'exit':
                break
                
            # display the 'Create Sheet' window if the user clicks the 'Create Sheet' button
            elif event == '-NEW-':
                window.hide()
                create_new_smartsheet_window()
                window.un_hide()

            # display the 'Import Data' window if the user clicks the 'Import Data' button
            elif event == '-IMPORT-':
                window.hide()
                import_sheet_window("Import", import_excel_data, event)                    
                window.un_hide()

            # display the 'Update Sheet' window if the user clicks the 'Update Sheet' button
            elif event == '-UPDATE-':
                window.hide()
                import_sheet_window("Update", import_excel_data, event)
                window.un_hide()

        # close the main GUI window
        window.close()
    except Exception as e:
        print(f"Error: {e}")
        print(traceback.format_exc())



if __name__ == "__main__":
    # Define the path to the configuration file
    SETTINGS_PATH = resource_path('core')

    # Load user settings from the configuration file
    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename=f'{SETTINGS_PATH}\\config.ini', use_config_file=True, convert_bools_and_none=True
    )

    # Get settings for the GUI's theme, font family, and font size
    theme = settings['GUI']['theme']
    font_family = settings['GUI']['font_family']
    font_size = int(settings['GUI']['font_size'])
    titlebar_font_size = int(settings['GUI']['titlebar_font_size'])
    
    # Get color settings for the GUI
    dark_blue = settings['COLORS']['dark_blue']
    lighter_blue = settings['COLORS']['lighter_blue']
    white = settings['COLORS']['white']
    french_gray = settings['COLORS']['french_gray']

    # Set the GUI's theme, font, and color settings
    sg.theme(theme)
    sg.set_options(
        background_color = dark_blue,
        border_width=0,
        button_color = (white, french_gray),
        element_background_color = (dark_blue),
        font = (font_family, font_size),
        icon='icon.ico',
        input_elements_background_color = (french_gray),
        input_text_color= dark_blue,
        text_element_background_color = (dark_blue),
        use_ttk_buttons=True,
    )

    # Run the main function to start the GUI
    main()