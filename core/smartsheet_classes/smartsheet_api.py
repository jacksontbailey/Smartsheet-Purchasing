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


    def add_rows(self, mapped_columns, column_dict):
        """Adds rows to the sheet.

        Args:
            rows_data (list): A list of dictionaries containing the data for each row.
            column_dict (dict): A dictionary of column names and IDs.

        Returns:
            list: A list of the added rows.
        """
        # - Create a list of row objects
        rows = []
        for mapped_column in mapped_columns:
            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            for column_name, value in mapped_column.items():
                column_id = column_dict[column_name]
                cell = smartsheet.models.Cell()
                cell.column_id = column_id
                cell.value = value
                new_row.cells.append(cell)
            rows.append(new_row)

        # - Add the rows to the sheet
        added_rows = self.smartsheet_client.Sheets.add_rows(self.sheet_id, rows)

        return added_rows


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


    def get_data(self):
        """Retrieves all the data from the sheet.

        Returns:
            pandas DataFrame: A DataFrame containing all the data from the sheet.
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


    def compare_data(self, excel_data):
        """
            Compares the provided excel data with the Smartsheet data and returns the differences between them.

            Args:
                excel_data (dict): The data to compare to the Smartsheet data. A dictionary with the item ID as key and the
                    sub-dictionary containing the data to compare as value.
                self.key_column_id (str): The name of the key representing the item ID.
                smartsheet_data (dict, optional): The Smartsheet data to compare with. If not provided, the sheet data will be
                    fetched from Smartsheet.

            Returns:
                dict: A dictionary containing the differences between the two datasets. The dictionary contains the item ID as
                    key and the sub-dictionary of differences as value. The sub-dictionary of differences contains only the keys
                    that exist in both datasets and have different values.
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
        """Updates the sheet data with the data to compare.

        Args:
            excel_data (dict): The data to compare to the sheet data.
            self.key_column_id (str): The name of the column containing the item IDs.
            smartsheet_data (dict, optional): The sheet data to compare. If not provided, the sheet data will be fetched from Smartsheet.
        """
        if self.smartsheet_data is None:
            self.smartsheet_data = self.get_data()

        differences = self.compare_data(excel_data)
        if not differences:
            return 'No differences found, nothing to update.'

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


        return 'Update successful.'


