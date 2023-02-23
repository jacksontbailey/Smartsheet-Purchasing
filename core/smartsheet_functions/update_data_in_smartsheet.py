import pandas as pd

def update_smartsheet(excel_data, sheet_manager):
    column_dict = sheet_manager.get_columns()

    # Create a list of dictionaries, where each dictionary represents a row in the smartsheet
    mapped_columns = [
        {
            "ITEM#": row["Item ID"],
            "ITEM DESCRIPTION": row["Item Description"],
            "QTY": row["Quantity"],
            "AREA": row["Area"],
            "NOTES": row["Specific Area"] if not pd.isna(row["Specific Area"]) else None,
        }
        for row in excel_data
    ]

    # Find the common keys between columns_to_compare and column_dict keys
    common_keys = set(mapped_columns[0].keys()) & set(column_dict.keys())

    # Get the column IDs for the common keys
    dictionary_of_keys = {key: column_dict[key] for key in common_keys}

    # Creating new dictionary where keys match the column ID in SmartSheet instead of the name
    new_dict = {row["ITEM#"]: {dictionary_of_keys[k]: v for k, v in row.items()} for row in mapped_columns}

    #key = dictionary_of_keys['ITEM#']
    sheet_manager.key_column_id = dictionary_of_keys['ITEM#']

    compared_data = sheet_manager.update_data(
        excel_data = new_dict, 
    )
    
    print(compared_data)