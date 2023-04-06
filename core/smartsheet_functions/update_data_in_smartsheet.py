import math

def update_smartsheet(excel_data, sheet_manager, option):
    """
    Updates data in a smartsheet based on an option.

    Parameters:
        excel_data (List[Dict]): A list of dictionaries containing the data to be updated.
        sheet_manager (Object): An object representing the smartsheet to update the data in.
        option (str): A string representing the update option, either "-IMPORT-" or anything else.

    Returns:
        str: A message indicating if the data was successfully updated or if duplicates were found.
    
    Raises:
    Exception: If there's an error updating the data in the smartsheet, an error message will be displayed.
    """
    try:
        mapped_columns=None
        column_dict = sheet_manager.get_columns()
        if option == "-IMPORT-":
            mapped_columns = [
                {
                    "ITEM#": row["Item ID"],
                    "ITEM DESCRIPTION": row["Item Description"],
                    "QTY": row["Quantity"],
                    "UOM": row["UOM"],
                    "AREA": row["Area"],
                    "NOTES": row["Specific Area"] if not math.isna(row["Specific Area"]) else "",
                    "AWARDED TO": row["Awarded To"],
                }
                for row in excel_data
            ]
        else:
            mapped_columns = [
                {
                    "ITEM#": row["Item ID"],
                    "ITEM DESCRIPTION": row["Item Description"],
                    "QTY": row["Quantity"],
                    "AREA": row["Area"],
                    "NOTES": row["Specific Area"] if not math.isna(row["Specific Area"]) else None,
                }
                for row in excel_data
            ]

            common_keys = set(mapped_columns[0].keys()) & set(column_dict.keys())
            dictionary_of_keys = {key: column_dict[key] for key in common_keys}
            sheet_manager.key_column_id = dictionary_of_keys['ITEM#']
            new_dict = {row["ITEM#"]: {dictionary_of_keys[k]: v for k, v in row.items()} for row in mapped_columns}

            if option == "-IMPORT-":
                added_excel_data = sheet_manager.add_rows(mapped_columns=mapped_columns, column_dict=column_dict, compare_dict_keys=new_dict)
                return added_excel_data if added_excel_data[0] == "Duplicates Found" else "Data inserted successfully!"
            else:
                compared_data = sheet_manager.update_data(excel_data=new_dict)
                return compared_data
    except Exception as e:
        return ["Incorrect Format"]