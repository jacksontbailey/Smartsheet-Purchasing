import PySimpleGUI as sg


def import_sheet_window(btn, import_function, option):
    # create the layout for the import sheet window
    layout = [
        [sg.T("What is the name of the SmartSheet you want to use?", expand_x=True, pad=((30, (20, 5))))], 
        [sg.Input(k='-SHEET NAME-', expand_x=True, pad=((30, (5, 10))))],
        [sg.T(text="Select an excel document to import to your Smartsheet.", pad=((30, (10, 5))))], 
        [sg.Input(k='-ATTACHMENT-', size=40, pad=((30, 5), (5, 20))), 
        sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),), size=10, pad=((5, 30), (5, 20)), change_submits=True)],
        [sg.Push(), sg.B(btn, border_width=0, k=btn.lower(), s=10, enable_events=True), sg.Cancel(tooltip='Exit the Application', s=10)],
    ]

    # create the import sheet window
    window = sg.Window('Import Data', layout, size=(450, 250))

    try:
    # event loop for the import sheet window
        while True:
            event, values = window.read()
            if event in (sg.WIN_CLOSED, 'Cancel'):
                window.close()
                break

            elif event == btn.lower():
                input_file_path = values['-ATTACHMENT-']
                selected_smartsheet_name = values['-SHEET NAME-']
                import_function(option, input_file_path, selected_smartsheet_name)
                window.close()

                if btn =="Update":
                    sg.popup_auto_close("Successfully updated your SmartSheet with the provided excel data!", title="Success", auto_close_duration=2)
                else:
                    sg.popup_auto_close("Successfully imported excel data into your SmartSheet!", title="Success" ,auto_close_duration=2)
                
    except Exception as e:
        sg.Print("Exception in my import sheet event loop for the program:", sg.__file__, e, keep_on_top=True, wait=True)
        sg.popup_error_with_traceback('Problem in my event loop!', e)
    
    window.close()