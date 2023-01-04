import smartsheet
from smartsheet_functions.client import create_smartsheet_client


class SheetManager:
    def __init__(self, api_key, sheet_id, sheet_data = None):
        """Initializes a SheetManager object.

        Args:
            sheet_id (int): The ID of the sheet to manage.
            sheet_data (pandas DataFrame, optional): The data for the sheet. If not provided, the data will be fetched from Smartsheet.
        """
        self.sheet_id = sheet_id
        self.sheet_data = sheet_data
        self.smartsheet_client = create_smartsheet_client(api_key)


    def get_columns(self):

        """Gets the column names and IDs for the sheet.

        Returns:
            dict: A dictionary of column names and IDs.
        """
        # - Get the sheet
        sheet = self.smartsheet_client.Sheets.get_sheet(self.sheet_id)

        # - Get the list of columns in the sheet
        columns = sheet.columns

        # - Create a dictionary of column names and IDs
        column_dict = {}
        for column in columns:
            column_dict[column.title] = column.id

        return column_dict


    def add_rows(self, rows_data, column_dict):
        """Adds rows to the sheet.

        Args:
            rows_data (list): A list of dictionaries containing the data for each row.
            column_dict (dict): A dictionary of column names and IDs.

        Returns:
            list: A list of the added rows.
        """
        # - Create a list of row objects
        rows = []
        for row_data in rows_data:
            new_row = smartsheet.models.Row()
            for column_name, value in row_data.items():
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


    def delete_rows(self, row_id):
        """Deletes a row from the sheet by its ID.

        Args:
            row_id (int): The ID of the row to delete.
        """
        # - Delete the row
        self.smartsheet_client.Sheets.delete_rows(self.sheet_id, row_id)


    def compare_data(self, data_to_compare, compare_column, sheet_data=None):
        """Compares the sheet data with the data to compare and returns the differences.

        Args:
            data_to_compare (pandas DataFrame): The data to compare to the sheet data.
            compare_column (str): The name of the column to compare.
            sheet_data (pandas DataFrame, optional): The sheet data to compare. If not provided, the sheet data will be fetched from Smartsheet.

        Returns:
            tuple: A tuple containing a boolean indicating whether there were differences and a dataframe of the differences.
        """
        if sheet_data is None:
            sheet_data = self.get_data()
            sheet_data_values = set(sheet_data[compare_column])
            data_to_compare_values = set(data_to_compare[compare_column])
            differences = sheet_data_values.symmetric_difference(data_to_compare_values)
            
            if differences:
                return (True, data_to_compare[data_to_compare[compare_column].isin(differences)])
            else:
                return (False, None)


    def update_data(self, data_to_compare, compare_column, sheet_data=None):
        """Updates the sheet data with the data to compare.
        
        Args:
            data_to_compare (pandas DataFrame): The data to compare to the sheet data.
            compare_column (str): The name of the column to compare.
            sheet_data (pandas DataFrame, optional): The sheet data to compare. If not provided, the sheet data will be fetched from Smartsheet.
        """
        if sheet_data is None:
            sheet_data = self.get_data()
            differences = self.compare_data(data_to_compare, compare_column, sheet_data)
            rows_to_update = []
            
            for index, row in sheet_data.iterrows():
                if row[compare_column] in differences[1][compare_column]:
                    compare_row = data_to_compare[data_to_compare[compare_column] == row[compare_column]]
                    row_to_update = {}
                    row_to_update["row_id"] = row["row_id"]
                    
                    for col in sheet_data.columns:
                        if row[col] != compare_row[col]:
                            row_to_update[col] = compare_row[col]
                    rows_to_update.append(row_to_update)
            
            self.update_rows(rows_to_update, self.get_columns())
