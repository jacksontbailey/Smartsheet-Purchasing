import os
import pandas as pd

class ExcelSheetManager:
    """Manages an Excel sheet.

    Args:
        file_name (str): The name of the Excel file.
        file_location (str): The location of the Excel file.
        tab_name (str): The name of the tab in the Excel file to be read.
        table_name (str, optional): The name of the table within the tab to be read.
    """
    
    def __init__(self, file_name = None, file_location=None, tab_name=None, filebrowse_url=None, table_name=None):
        self.file_name = file_name
        self.file_location = file_location
        self.tab_name = tab_name
        self.filebrowse_url = filebrowse_url
        self.table_name = table_name
        self.data = None



    def check_file(self):
        """Checks if the Excel file is open or not.

        Returns:
            bool: True if the file is closed and can be read, False otherwise.
            str: A message to the user if the file is open or not found.
        """

        try:
            # Checks if the Excel file is open
            if os.path.isfile(fr"{self.file_location}/{self.file_name}.xlsx"):
                # Tries to open the file in read-only mode
                with open(fr"{self.file_location}/{self.file_name}.xlsx", "r") as f:
                    pass
            elif os.path.isfile(fr"{self.filebrowse_url}"):
                with open(fr"{self.filebrowse_url}", "r") as f:
                    pass
            else:
                # Returns False and an error message
                return False, "File not found"

            # Returns True and no message
            return True, None

        except PermissionError:
            # Returns False and a message to the user to close the file first
            return False, "Please close the Excel file before reading it"



    def check_tab(self):
        """Checks if the tab name exists in the Excel file.

        Returns:
            bool: True if the tab name is valid, False otherwise.
            str: A message to the user if the tab name is invalid.
        """

        # Creates an ExcelFile object from the Excel file
        # Uses a ternary operator to choose the file path based on the filebrowse_url attribute
        xls = pd.ExcelFile(self.filebrowse_url if self.filebrowse_url else fr"{self.file_location}/{self.file_name}.xlsx")

        # Checks if the tab name is in the list of sheet names
        if self.tab_name in xls.sheet_names:
            # Returns True and no message
            return True, None
        else:
            # Returns False and an error message
            return False, "Incorrect tab name"
        


    def read_data(self):
        """Reads data from the specified tab in the Excel file.

        Returns:
            List[Dict[str, Any]]: A list of dictionaries, where each dictionary represents a row in the tab and contains data for each column.
            Dict[str, int]: A dictionary that contains the names and IDs of the columns in the tab.
            str: A message to the user if there is an error in reading the data.
        """

        # Calls the check_file and check_tab methods to see if the file and tab can be read
        # Uses a for loop to iterate over the methods and check their status
        for check_method in [self.check_file, self.check_tab]:
            status, message = check_method()

            # If the status is False, returns the message
            if not status:
                return message

        try:
            # Reads data from the specified tab in the Excel file
            # Uses a ternary operator to choose the file path based on the filebrowse_url attribute
            df = pd.read_excel(self.filebrowse_url if self.filebrowse_url else fr"{self.file_location}/{self.file_name}.xlsx", engine='openpyxl', sheet_name=self.tab_name, header=0)

            # Checks if the Excel tab is empty
            if df.empty:
                return 

            self.data = df.to_dict('records')
            return self.data

        except Exception as e:
            return f"An error occurred while reading the data: {e}"
