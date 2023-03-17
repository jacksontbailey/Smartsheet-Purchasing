from core.smartsheet_functions.create_new_sheet import create_new_smartsheet
import PySimpleGUI as sg

def create_new_smartsheet_window():
    # create the layout for the new SmartSheet window
    layout = [
        [sg.Text('What do you want to call your new SmartSheet?')],
        [sg.Input(key='new_sheet_name')],
        [sg.Button('Create', key='create')]
    ]

    # create the new SmartSheet window
    new_sheet_window = sg.Window('New SmartSheet', layout)

    # event loop for the new SmartSheet window
    while True:
        event, values = new_sheet_window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'create':
            new_sheet_name = values['new_sheet_name']
            create_new_smartsheet(new_sheet_name)
            new_sheet_window.close()
            return