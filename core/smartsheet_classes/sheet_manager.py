import smartsheet
from smartsheet_functions.client import create_smartsheet_client


class SheetManager:
    def __init__(self, sheet_id):
        self.sheet_id = sheet_id
        self.smartsheet_client = create_smartsheet_client()


    def get_columns(self):
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
        # - Get the row
        row = self.smartsheet_client.Sheets.get_row(self.sheet_id, row_id)

        return row

    def delete_rows(self, row_id):
        # - Delete the row
        self.smartsheet_client.Sheets.delete_rows(self.sheet_id, row_id)
