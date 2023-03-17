import PySimpleGUI as sg

def select_smartsheet_window():
    # create the layout for the Select SmartSheet window
    layout = [
        [sg.Text('What is the name of the SmartSheet you want to use?')],
        [sg.Input(key='sheet_name')],
        [sg.Button('Submit', key='submit')]
    ]

    # create the new SmartSheet window
    select_sheet_window = sg.Window('New SmartSheet', layout)

    # event loop for the new SmartSheet window
    while True:
        event, values = select_sheet_window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'submit':
            sheet_name = values['sheet_name']
            select_sheet_window.close()
            return sheet_name