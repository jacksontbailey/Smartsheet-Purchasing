import PySimpleGUI as sg


def import_sheet_window(btn, import_function, option):
    """
    Display an import sheet window and process the user input for importing or updating data in Smartsheets.
    Args:
        btn (str): The button name indicating whether the user wants to import or update data in Smartsheets.
        import_function (function): The function that performs the import or update operation.
        option (str): The option indicating whether the import is a regular or update operation.

    Returns:
        None

    Raises:
        Exception: If there is an error in the event loop for the import sheet window.
    """
    # create the layout for the import sheet window
    frame_layout = [
        [
            sg.T(
                "What is the name of the Smartsheet you want to use?",
                expand_x=True,
                pad=((30, (10, 5))),
            )
        ],
        [sg.Input(k="-SHEET NAME-", expand_x=True, pad=((30, (5, 10))))],
        [
            sg.T(
                text="Select an excel document to import to your Smartsheet.",
                pad=((30, (10, 5))),
            )
        ],
        [
            sg.Input(k="-ATTACHMENT-", size=40, pad=((30, 10), (5, 20))),
            sg.FileBrowse(
                file_types=(("Excel Files", "*.xls*"),),
                font=("any 10 bold"),
                size=10,
                pad=((5, 30), (5, 20)),
                change_submits=True,
            ),
        ],
    ]

    layout = [
        [sg.Frame(f"Excel File: {btn}", frame_layout, font="Any 16", expand_x=True)],
        [
            sg.Push(),
            sg.B(
                btn,
                k=btn.lower(),
                font=("any 10 bold"),
                enable_events=True,
                pad=((0, 5), (10, 0)),
                size=(25, 1)
            ),
            sg.Cancel(s=15, font=("any 10 bold"), pad=((5, 5), (10, 0))),
        ],
    ]
    # create the import sheet window
    window = sg.Window(
        f"Smartsheet Data: {btn}", layout, finalize=True, size=(450, 250)
    )

    window["Browse"].set_cursor(cursor="hand2")
    window[f"{btn.lower()}"].set_cursor(cursor="hand2")
    window["Cancel"].set_cursor(cursor="hand2")

    # event loop for the import sheet window
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Cancel"):
            window.close()
            break

        elif event == btn.lower():
            try:
                input_file_path = values["-ATTACHMENT-"]
                selected_smartsheet_name = values["-SHEET NAME-"]

                if not input_file_path or not selected_smartsheet_name:
                    window.disappear()
                    sg.popup_error("Form is missing data.")
                    window.reappear()
                    break

                import_function_results = import_function(
                    option, input_file_path, selected_smartsheet_name
                )
                error_list = None
                import_key = import_function_results[0]


                if len(import_function_results) == 2:
                    error_list = ", ".join(import_function_results[1])

                error_messages = {
                    "Duplicates Found": f"Error: Could not import data from the excel file to smartsheets due to matching Item IDs in both sheets.\n\nDuplicate Item IDs: \n\n{error_list}",
                    "Empty Tab": f"Error: The Purchasing_Items tab in your excel document is empty.",
                    "Incorrect Format": f"Error: The excel file you provided could not be imported due to having the incorrect format.",
                    "Incorrect Tab Name": f"Error: \n\nPlease ensure that the primary tab in your attached excel file is named 'Purchasing_Items'. Other tabs will not be read. \n\nExcel file you uploaded: \n\n'{input_file_path}'",
                    "Invalid ID": f"Smartsheet named '{selected_smartsheet_name}' could not be found. Please make sure you type in the exact name of the Smartsheet.",
                    "No Differences": f"There were no differences found when comparing the data in the excel file and the smartsheet. \n\nExcel file you uploaded: \n\n'{input_file_path}'",
                }

                if import_key in error_messages:
                    window.disappear()
                    sg.popup_ok(error_messages.get(import_key), title=import_key)
                    window.reappear()
                    break

                window.close()

                if btn == "Update":
                    sg.popup_auto_close(
                        "Successfully updated your Smartsheet with the provided excel data!",
                        title="Success",
                        auto_close_duration=2,
                    )
                else:
                    sg.popup_auto_close(
                        "Successfully imported excel data into your Smartsheet!",
                        title="Success",
                        auto_close_duration=2,
                    )
            except Exception as e:
                sg.Print(
                    "Exception in my import sheet event loop for the program:",
                    sg.__file__,
                    e,
                    keep_on_top=True,
                    wait=True,
                )
                sg.popup_error_with_traceback("Problem in my event loop!")

    window.close()
