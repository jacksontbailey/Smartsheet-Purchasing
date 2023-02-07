import pandas as pd

class ExcelSheetManager:
    """Manages an Excel sheet.

    Args:
        file_name (str): The name of the Excel file.
        file_location (str): The location of the Excel file.
        tab_name (str): The name of the tab in the Excel file to be read.
        table_name (str, optional): The name of the table within the tab to be read.
    """
    
    def __init__(self, file_name, file_location, tab_name, table_name=None):
        self.file_name = file_name
        self.file_location = file_location
        self.tab_name = tab_name
        self.table_name = table_name
        self.data = None

    def read_data(self):
        """Reads data from the specified tab in the Excel file.

        Returns:
            List[Dict[str, Any]]: A list of dictionaries, where each dictionary represents a row in the tab and contains data for each column.
            Dict[str, int]: A dictionary that contains the names and IDs of the columns in the tab.
        """
        try:
            if self.table_name:
                df = pd.read_excel(f"{self.file_location}\{self.file_name}.xlsx", sheet_name=self.tab_name, engine='openpyxl', header=None)
                table = df.iloc[df[df[0].str.contains(self.table_name, na=False)].index[0] + 1:]
                table = table[:table[table.isna().all(1)].index[0]]
                self.data = table.iloc[1:].reset_index(drop=True).to_dict('records')
                column_names_ids = {column: index for index, column in enumerate(table.iloc[0].tolist())}
            else:
                df = pd.read_excel(f"{self.file_location}/{self.file_name}", sheet_name=self.tab_name)
                self.data = df.to_dict('records')
                column_names_ids = {column: index for index, column in enumerate(df.columns)}

            return self.data, column_names_ids
        except Exception as e:
            return str(e)
