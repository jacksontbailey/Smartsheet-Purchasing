from core.smartsheet_functions.create_new_sheet import create_new_smartsheet
import PySimpleGUI as sg

def create_new_smartsheet_window():
    # create the layout for the new SmartSheet window
    layout = [
        [sg.T('What do you want to call your new SmartSheet?')],
        [sg.Input(key='-IN-')],
        [sg.B('Create', border_width=0, key='-CREATE-'), sg.Cancel()]
    ]

    # create the new SmartSheet window
    window = sg.Window('New SmartSheet', layout, finalize=True)
    
    window['-CREATE-'].set_cursor(cursor="hand2")
    window['Cancel'].set_cursor(cursor="hand2")

    try:
        # event loop for the new SmartSheet window
        while True:
            event, values = window.read()
            if event == sg.WINDOW_CLOSED or event == 'Cancel':
                window.close()
                break
            elif event == '-CREATE-':
                new_sheet_name = values['-IN-']
                create_new_smartsheet(new_sheet_name)
                window.close()
                sg.popup_auto_close(title="Successfully created a new SmartSheet!", auto_close_duration=2)
                return
    
    except Exception as e:
        sg.Print("Exception in my import sheet event loop for the program:", sg.__file__, e, keep_on_top=True, wait=True)
        sg.popup_error_with_traceback('Problem in my event loop!', e)
    
    window.close()