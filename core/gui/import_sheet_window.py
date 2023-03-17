from core.gui.select_smartsheet_window import select_smartsheet_window
import PySimpleGUI as sg


def import_sheet_window(btn, import_function, option):
    # create the layout for the import sheet window
    layout = [
        [sg.Input(key='sheet_file'), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),))],
        [sg.Button(btn, key=btn.lower())]
    ]

    # create the import sheet window
    import_window = sg.Window('Import Sheet', layout)

    # event loop for the import sheet window
    while True:
        event, values = import_window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == btn.lower():
            input_file_path = values['sheet_file']
            selected_smartsheet_name = select_smartsheet_window()
            import_function(option, input_file_path, selected_smartsheet_name)
            import_window.close()
            return