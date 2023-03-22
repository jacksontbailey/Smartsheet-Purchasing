import PySimpleGUI as sg


def import_sheet_window(btn, import_function, option):
    # create the layout for the import sheet window
    frame_layout = [
        [sg.T("What is the name of the Smartsheet you want to use?", background_color='#002852', expand_x=True, pad=((30, (10, 5))))], 
        [sg.Input(k='-SHEET NAME-', expand_x=True, pad=((30, (5, 10))))],
        [sg.T(text="Select an excel document to import to your Smartsheet.", background_color='#002852', pad=((30, (10, 5))))], 
        [sg.Input(k='-ATTACHMENT-', size=40, pad=((30, 5), (5, 20))), 
        sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),), size=10, pad=((5, 30), (5, 20)), change_submits=True)],
    ]

    layout = [
        [sg.Frame(f'{btn} Excel Data', frame_layout, background_color='#002852', font='Any 16', title_color='white', expand_x=True)],
        [sg.Push(background_color='#002852'), sg.B(btn, k=btn.lower(), s=10, enable_events=True, pad=((0, 5), (10, 0))), sg.Cancel(s=10, pad=((5, 5), (10, 0)))],
    ]
    # create the import sheet window
    window = sg.Window(f'Smartsheet Data: {btn}', layout, background_color='#002852', finalize=True, size=(450, 250))

    window['Browse'].set_cursor(cursor="hand2")
    window[f'{btn.lower()}'].set_cursor(cursor="hand2")
    window['Cancel'].set_cursor(cursor="hand2")


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
                
                if not input_file_path or not selected_smartsheet_name:
                    sg.popup_error("Form is missing data.", background_color='#002852')
                    window.close()
                    break

                import_function(option, input_file_path, selected_smartsheet_name)
                window.close()

                if btn =="Update":
                    sg.popup_auto_close("Successfully updated your Smartsheet with the provided excel data!", title="Success", auto_close_duration=2)
                else:
                    sg.popup_auto_close("Successfully imported excel data into your Smartsheet!", title="Success" ,auto_close_duration=2)

    except Exception as e:
        sg.Print("Exception in my import sheet event loop for the program:", sg.__file__, e, keep_on_top=True, wait=True)
        sg.popup_error_with_traceback('Problem in my event loop!', e)
    
    window.close()