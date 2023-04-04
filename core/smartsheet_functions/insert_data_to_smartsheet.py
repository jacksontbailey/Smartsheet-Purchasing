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
    mapped_columns = [
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
        # Find the common keys between columns_to_compare and column_dict keys
        common_keys = set(mapped_columns[0].keys()) & set(column_dict.keys())

        # Get the column IDs for the common keys
        dictionary_of_keys = {key: column_dict[key] for key in common_keys}
        #key = dictionary_of_keys['ITEM#']
        sheet_manager.key_column_id = dictionary_of_keys['ITEM#']
        # Creating new dictionary where keys match the column ID in SmartSheet instead of the name
        new_dict = {row["ITEM#"]: {dictionary_of_keys[k]: v for k, v in row.items()} for row in mapped_columns}


        # Insert the rows into the smartsheet
        added_excel_data = sheet_manager.add_rows(mapped_columns = mapped_columns, column_dict = column_dict, compare_dict_keys = new_dict)
        
        print(f"added_excel_data is {added_excel_data}, type: {type(added_excel_data)}")
        if added_excel_data[0] == "Duplicates Found":
            return added_excel_data
        
        return "Data inserted successfully!"
    except Exception as e:
        # If there's an error, display an error message
        print("Error inserting data: ", str(e))
        return("Error")


# Note: The columns for the data being inserted (ITEM#, ITEM DESCRIPTION, QTY, UOM, AREA, NOTES, AWARDED TO)
# -- are specific to this project and need to be adjusted on a project-to-project basis.

