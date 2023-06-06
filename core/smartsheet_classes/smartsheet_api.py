import smartsheet
import pandas as pd
from ..smartsheet_functions.client import create_smartsheet_client


class SmartSheetApi:
    def __init__(self, api_key, folder_id = None, new_sheet_name = None, sheet_data = None, sheet_id = None, sheet_name = None, workspace_id = None):
        """Initializes a SheetManager object.

        Args:
            sheet_id (int): The ID of the sheet to manage.
            sheet_data (pandas DataFrame, optional): The data for the sheet. If not provided, the data will be fetched from Smartsheet.
        """
        self.folder_id = folder_id
        self.key_column_id = None
        self.new_sheet_name = new_sheet_name
        self.sheet_columns = None
        self.sheet_data = sheet_data
        self.sheet_id = sheet_id
        self.sheet_name = sheet_name
        self.smartsheet_client = create_smartsheet_client(api_key)
        self.smartsheet_data = None
        self.workspace_id = workspace_id


    def get_sheets_in_folder(self):
        """Returns the names and IDs of all sheets within a folder or workspace.
        
        Returns:
            list: A list of dictionaries, where each dictionary contains the name and ID of a sheet.
        """
        sheets = self.smartsheet_client.Folders.list_sheets(self.folder_id).result.data
        return [{"name": sheet.name, "id": sheet.id} for sheet in sheets]
    

    def get_sheets_in_workspace(self):
        """Returns the names and IDs of all sheets within a workspace.
        
        Returns:
            list: A list of dictionaries, where each dictionary contains the name and ID of a sheet.
        """
        sheets = self.smartsheet_client.Workspaces.get_workspace(self.workspace_id).sheets
        return [{"name": sheet.name, "id": sheet.id} for sheet in sheets]


    def get_sheet_id_by_name(self):
        """Returns the ID of a sheet by its name.
        
        Returns:
            int: The ID of the sheet.
        """
        try:
            if not hasattr(self, 'sheets'):
                if self.folder_id:
                    print("has folder id")
                    self.sheets = self.get_sheets_in_folder()
                elif self.workspace_id:
                    print("has workspace id")
                    self.sheets = self.get_sheets_in_workspace()
                else: 
                    print("The SmartSheetApi objects, folder_id and workspace_id, have no value. Please re-run after correcting issue.")
                    return
            
            sheet = next((sheet for sheet in self.sheets if sheet["name"] == self.sheet_name), None)
            if sheet:
                self.sheet_id = sheet["id"]
                return sheet["id"] 
            else:
                None

        except Exception as e:
            print(str(e))



    def create_sheet_in_workspace(self, template_sheet_id, new_sheet_name,):
        """Creates a new sheet in a workspace from a template.
        
        Args:
            template_sheet_id (int): The ID of the template sheet to use.
            new_sheet_name (str): The name of the new sheet.
        
        Returns:
            Sheet: The new sheet that was created.
        """
        # - Create a new sheet from the template
        new_sheet = self.smartsheet_client.Workspaces.create_sheet_in_workspace_from_template(
            self.workspace_id,
            smartsheet.models.Sheet({
                'name': new_sheet_name,
                'from_id': int(template_sheet_id)
            })
        ).result

        # - Return the new sheet
        return new_sheet
    
    
    def create_sheet_in_folder(self, template_sheet_id, new_sheet_name):
        """Creates a new sheet in a folder from a template.
        
        Args:
            template_sheet_id (int): The ID of the template sheet to use.
            new_sheet_name (str): The name of the new sheet.
        
        Returns:
            Sheet: The new sheet that was created.
        """
        new_sheet = self.smartsheet_client.Folders.create_sheet_from_template(
            self.folder_id,
            smartsheet.models.Sheet({
                'name': new_sheet_name,
                'from_id': int(template_sheet_id)
            })
        ).result

        return new_sheet


    def get_columns(self):

        """Gets the column names and IDs for the sheet.

        Returns:
            dict: A dictionary of column names and IDs.
        """
        # - Get the sheet
        sheet = self.smartsheet_client.Sheets.get_sheet(self.sheet_id)

        # - Get the list of columns in the sheet
        all_sheet_columns = sheet.columns

        # - Create a dictionary of column names and IDs
        self.sheet_columns = {}
        for column in all_sheet_columns:
            self.sheet_columns[column.title] = column.id

        return self.sheet_columns


    def update_columns(self, sheet_id, new_columns):
        """Updates the columns of a sheet.
        
        Args:
            sheet_id (int): The ID of the sheet to update.
            new_columns (list): A list of dictionaries, where each dictionary represents a column and contains the following keys:
                - title (str): The title of the column.
                - type (str): The type of the column. Possible values are: 
                  "TEXT_NUMBER", "DATE", "CONTACT_LIST", "CHECKBOX", "DURATION", "PERCENTAGE".
                - options (dict, optional): A dictionary of options for the column, such as width or validation rules.
        Returns:
            Sheet: The updated sheet.
        """
        sheet = self.smartsheet_client.Sheets.get(sheet_id).result
        existing_columns = {col.title: col for col in sheet.columns}

        for col in new_columns:
            if col["title"] in existing_columns:
                updated_column = existing_columns[col["title"]]
                updated_column.type = col["type"]
                if "options" in col:
                    updated_column.options = col["options"]
            else:
                sheet.columns.append(smartsheet.models.Column(col))

        self.smartsheet_client.Sheets.update(sheet)
        return self.smartsheet_client.Sheets.get(sheet_id).result


    def get_row(self, row_id):
        """Gets a row from the sheet by its ID.

        Args:
            row_id (int): The ID of the row to retrieve.

        Returns:
            Row: The row object with the specified ID.
        """
        # - Get the row
        row = self.smartsheet_client.Sheets.get_row(self.sheet_id, row_id)

        return row


    def add_rows(self, mapped_columns, column_dict, compare_dict_keys, allowed):
        """Adds rows to the sheet.

        Args:
            rows_data (list): A list of dictionaries containing the data for each row.
            column_dict (dict): A dictionary of column names and IDs.
            allowed (bool): True if duplicates are allowed, False otherwise.

        Returns:
            list: A list of the added rows.
        """
        # - Create a list of row objects
        rows = []

        duplicates = self.check_duplicates(compare_dict_keys)
        print(f"duplicates are: {duplicates}")

        if duplicates and not allowed:
            return ["Duplicates Found", duplicates]
        
        for mapped_column in mapped_columns:
            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            for column_name, value in mapped_column.items():
                column_id = column_dict[column_name]
                cell = smartsheet.models.Cell()
                cell.column_id = column_id
                cell.value = f"{value} - duplicate" if value in duplicates and allowed else value
                print(f'cell value is: {cell.value}')
                new_row.cells.append(cell)
            rows.append(new_row)

        self.smartsheet_client.Sheets.add_rows(self.sheet_id, rows)
        
        return "Success"


    def update_rows(self, rows_data, column_dict):
        """Updates rows in the sheet with the provided data.
    
        Args:
            rows_data (list): A list of dictionaries, where each dictionary contains the row data to update.\
                                The dictionary should include the column name as the key and the new value as the value.\
                                It should also include a "row_id" key with the ID of the row to update.
            column_dict (dict): A dictionary of column names and IDs.
    
        Returns:
            list: A list of Row objects for the updated rows.
        """
        # - Create a list of row objects
        rows = []
        for row_data in rows_data:
            row_id = row_data["row_id"]
            row = self.smartsheet_client.Sheets.get_row(self.sheet_id, row_id)
            for column_name, value in row_data.items():
                if column_name == "row_id":
                    continue
                column_id = column_dict[column_name]
                cell = [c for c in row.cells if c.column_id == column_id][0]
                cell.value = value
            rows.append(row)

        # - Update the rows
        updated_rows = self.smartsheet_client.Sheets.update_rows(self.sheet_id, rows)

        return updated_rows


    def delete_rows(self, row_id):
        """Deletes a row from the sheet by its ID.

        Args:
            row_id (int): The ID of the row to delete.
        """
        # - Delete the row
        self.smartsheet_client.Sheets.delete_rows(self.sheet_id, row_id)


    def highlight_duplicates(self, duplicate_ids):
        """Highlights all rows with duplicate IDs.

        Args:
        duplicate_ids (list): A list of IDs for rows with duplicates.
        """
        print(f"duplicates are {duplicate_ids}")
        format_string = ",,,,,,,,,25,,,,,,,"
        for duplicate_id in duplicate_ids:
            row = self.smartsheet_client.Sheets.get_row(self.sheet_id, duplicate_id)
            for cell in row.cells:
                cell.format_ = smartsheet.models.Format()
                cell.format_.background_color = format_string

        self.smartsheet_client.Sheets.update_rows(self.sheet_id, [row])



    def get_data(self):
        """
        Retrieves the data from the sheet based on the specified column IDs.

        Args:
            column_ids (List[str]): A list of column IDs to retrieve data from.

        Returns:
            dict: A dictionary representing the sheet data, where the keys are the row IDs and the values are dictionaries representing the column values for each row.

        Raises:
            ValueError: If the provided column IDs are not a list.
            ValueError: If any of the provided column IDs are not strings.
            Exception: If there is an error retrieving data from Smartsheet.
        """
        # - Fetch the sheet data
        columns_ids = self.sheet_columns.values()
        rows = self.smartsheet_client.Sheets.get_sheet(self.sheet_id).rows
        # - Build the data frame
        data = {}
        for row in rows:
            row_data = {}
            has_data = False
            
            for cell in row.cells:
                column = next(c for c in columns_ids if c == cell.column_id)
                row_data['row'] = row.id
                row_data[column] = cell.value
                has_data = True
            if has_data:
                item_id = row_data.get(self.key_column_id)
                
                if item_id is not None:
                    data[item_id] = row_data
                else:
                    row_data.clear()
        return data

    def check_duplicates(self, compare_dict_keys):
        """
        Checks if there are any duplicate values in the specified column.

        Args:
            compare_dict_keys (str): The ID of the column to check for duplicates.

        Returns:
            bool: True if there are duplicates in the column, False otherwise.

        Raises:
            ValueError: If the provided column ID is not a string.
            KeyError: If the provided column ID is not found in the sheet data.
        """
        if self.smartsheet_data is None:
            self.smartsheet_data = self.get_data()

        num_rows = len(self.smartsheet_client.Sheets.get_sheet(self.sheet_id).rows)
        # - If the sheet is empty, return None
        if num_rows == 0:
            return None


        duplicates = []

        for key in compare_dict_keys.keys() & self.smartsheet_data.keys():
            duplicates.append(key)

        if not duplicates:
            return None
        else:
            return duplicates


    def compare_data(self, excel_data):
        """
        Compares the Excel data with the Smartsheet data based on the specified key column ID.

        Args:
            excel_data (dict): A dictionary representing the Excel data to compare.
            smartsheet_data (dict): A dictionary representing the Smartsheet data to compare.
            key_column_id (str): The ID of the key column to compare.

        Returns:
            Tuple[List[dict], List[dict], List[dict]]: A tuple of three lists representing the rows to add, update, and delete from the Smartsheet data.

        Raises:
            ValueError: If the provided Excel data is not a dictionary.
            ValueError: If the provided Smartsheet data is not a dictionary.
            ValueError: If the provided key column ID is not a string.
            KeyError: If the provided key column ID is not found in either the Excel or Smartsheet data.
        """
        if self.smartsheet_data is None:
            self.smartsheet_data = self.get_data()

        differences = {}

        for key in excel_data.keys() & self.smartsheet_data.keys():
            subdict1 = self.smartsheet_data[key]
            subdict2 = excel_data[key]
            subdict_diff = {}
            
            # Compare sub-dictionaries
            for subkey in subdict1.keys() & subdict2.keys():
                if subdict1[subkey] != subdict2[subkey]:
                    subdict_diff[subkey] = (subdict1[subkey], subdict2[subkey])

            if subdict_diff:
                differences[key] = subdict_diff
        return differences


    def update_data(self, excel_data):
        """
        Updates the sheet data with the data to compare.

        Args:
            excel_data (dict): A dictionary representing the Excel data to be compared with the Smartsheet data.
            key_column_id (str): The name of the column in both the Excel and Smartsheet data that contains the unique identifier for each item.
            smartsheet_data (dict, optional): A dictionary representing the Smartsheet data to be compared with the Excel data. If not provided, the Smartsheet data will be fetched using the Smartsheet API.

        Raises:
            ValueError: If the provided Excel data is not a dictionary.
            ValueError: If the provided key column ID is not a string.
            ValueError: If the provided Smartsheet data is not a dictionary.
            KeyError: If the provided key column ID is not found in either the Excel or Smartsheet data.
            Exception: If there is an error updating the Smartsheet data with the Excel data.
        """
        if self.smartsheet_data is None:
            self.smartsheet_data = self.get_data()

        differences = self.compare_data(excel_data)
        print(f"dif: {differences}")
        if not differences:
            print("NONE")
            return ['No Differences']

        rows_to_update = []
        for row_id, row_changes in differences.items():
            row = self.smartsheet_data[row_id].get("row")
            cells_to_update = []
            for col_id, (old_value, new_value) in row_changes.items():
                new_cell = self.smartsheet_client.models.Cell()
                new_cell.column_id = col_id
                new_cell.value = new_value
                cells_to_update.append(new_cell)
            
            new_row = self.smartsheet_client.models.Row()
            new_row.id = row
            new_row.cells = cells_to_update
            rows_to_update.append(new_row)

        self.smartsheet_client.Sheets.update_rows(
            self.sheet_id, 
            rows_to_update        
        )

        return ['Update successful.']


