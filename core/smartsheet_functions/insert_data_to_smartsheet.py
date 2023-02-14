import pandas as pd

def insert_data_to_smartsheet(excel_data, sheet_manager):
    """
    Inserts data into a smartsheet.

    Parameters:
    excel_data (List[Dict]): A list of dictionaries containing the data to be inserted.
    sheet_manager (Object): An object representing the smartsheet to insert the data into.

    Returns:
    None
    
    Raises:
    Exception: If there's an error inserting the data into the smartsheet, an error message will be displayed.
    """

    # Get the columns of the smartsheet
    column_dict = sheet_manager.get_columns()

    # Create a list of dictionaries, where each dictionary represents a row in the smartsheet
    rows_data = [
        {
            "ITEM#": row["Item ID"],
            "ITEM DESCRIPTION": row["Item Description"],
            "QTY": row["Quantity"],
            "UOM": row["UOM"],
            "AREA": row["Area"],
            "NOTES": row["Specific Area"] if not pd.isna(row["Specific Area"]) else "",
            "AWARDED TO": row["Awarded To"],
        }
        for row in excel_data
    ]

    try:
        # Insert the rows into the smartsheet
        added_rows = sheet_manager.add_rows(rows_data, column_dict)
        print("Data inserted successfully!")
        print(added_rows)
    except Exception as e:
        # If there's an error, display an error message
        print("Error inserting data: ", str(e))


# Note: The columns for the data being inserted (ITEM#, ITEM DESCRIPTION, QTY, UOM, AREA, NOTES, AWARDED TO)
# -- are specific to this project and need to be adjusted on a project-to-project basis.

