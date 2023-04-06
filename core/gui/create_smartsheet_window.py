from ..smartsheet_functions.create_new_sheet import create_new_smartsheet
import PySimpleGUI as sg

def create_new_smartsheet_window():
    """
    Display a window for creating a new Smartsheet and process the user input for creating a new sheet.

    Returns:
        None

    Raises:
        Exception: If there is an error in the event loop for the new Smartsheet window.
    """
    # create the layout for the new Smartsheet window
    frame_layout = [
        [sg.T('What do you want to call your new Smartsheet?', expand_x=True, pad=((30, (20, 5))))],
        [sg.Input(key='-IN-', expand_x=True, pad=((30, (5, 20))))],
    ]

    layout = [
        [sg.Frame('Create a Smartsheet', frame_layout, font='Any 16', expand_x=True)],
        
        [
            sg.Push(), 
            sg.B('Create', font=('any 10 bold' ), k='-CREATE-', s=10, enable_events=True, pad=((0, 5), (20, 0))), 
            sg.Cancel(font=('any 10 bold' ), s=10, pad=((5, 5), (20, 0)))
        ],
    ]

    # create the new Smartsheet window
    window = sg.Window('New Smartsheet', layout, finalize=True, size=(400, 200))
    
    window['-CREATE-'].set_cursor(cursor="hand2")
    window['Cancel'].set_cursor(cursor="hand2")

    try:
        # event loop for the new Smartsheet window
        while True:
            event, values = window.read()
            if event == sg.WINDOW_CLOSED or event == 'Cancel':
                window.close()
                break
            elif event == '-CREATE-':
                new_sheet_name = values['-IN-']
                
                if not new_sheet_name:
                    sg.popup_error("Form is missing data.")
                    window.close()
                    break

                create_new_smartsheet(new_sheet_name)
                window.close()
                sg.popup_auto_close("Successfully created a new Smartsheet!", title="Success", auto_close_duration=2)
                return
    
    except Exception as e:
        sg.Print("Exception in my import sheet event loop for the program:", sg.__file__, e, keep_on_top=True, wait=True)
        sg.popup_error_with_traceback('Problem in my event loop!', e)
    
    window.close()