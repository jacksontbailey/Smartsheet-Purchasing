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



from core.gui.create_smartsheet_window import create_new_smartsheet_window
from core.gui.import_sheet_window import import_sheet_window
from core.smartsheet_functions.import_excel_data import import_excel_data
import PySimpleGUI as sg

def main():
    sg.theme("DarkTeal6")
    # define the layout for the main GUI window

    frame_layout = [
        [sg.Button("Import Data", border_width=0, k='-IMPORT-', tooltip='Import excel data into an empty smartsheet', button_color=('white', '#C4961B'), expand_x=True, p=((30, 7.5), (20, 10))),
         sg.Button("Update Sheet", border_width=0, k='-UPDATE-', tooltip='Update the existing smartsheet with an excel file', expand_x=True, p=((7.5, 30), (20, 10)))],
        [sg.Button('Create Sheet', border_width=0, k='-NEW-', tooltip='Create a new smartsheet', expand_x=True, p=(30,(10,20)))],
    ]

    layout = [
        [sg.Frame('Smartsheet Actions', frame_layout, font='Any 16', title_color='white', expand_x=True)],
        [sg.Push(), sg.Cancel(button_text="Exit", k='exit', s=10)]
    ]

    # create the main GUI window
    window = sg.Window('PSC - Purchasing Smartsheet Creator', layout, size=(400, 200), finalize=True)

    window['-NEW-'].set_cursor(cursor="hand2")
    window['-UPDATE-'].set_cursor(cursor="hand2")
    window['-IMPORT-'].set_cursor(cursor="hand2")
    window['exit'].set_cursor(cursor="hand2")

    # event loop for the main GUI window
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'exit':
            break
        elif event == '-NEW-':
            window.hide()
            create_new_smartsheet_window()
            window.un_hide()

        elif event == '-IMPORT-':
            window.hide()
            import_sheet_window("Import", import_excel_data, event)
            window.un_hide()

        elif event == '-UPDATE-':
            window.hide()
            import_sheet_window("Update", import_excel_data, event)
            window.un_hide()

    # close the main GUI window
    window.close()

# run the main GUI window
if __name__ == "__main__":
    main()