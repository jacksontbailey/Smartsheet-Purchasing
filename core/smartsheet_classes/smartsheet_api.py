import smartsheet
from smartsheet_functions.client import create_smartsheet_client


class SmartSheetApi:
    """A class for accessing Smartsheet data using the Smartsheet API.
    
    Args:
        api_key (str): The API key to use for authentication.
        folder_or_workspace_id (str, optional): The ID of a folder or workspace.
        sheet_name (str, optional): The name of a sheet.
    """

    def __init__(self, api_key, folder_or_workspace_id=None, sheet_name=None):
        self.smartsheet_client = create_smartsheet_client(api_key)
        self.folder_or_workspace_id = folder_or_workspace_id
        self.sheet_name = sheet_name


    def get_sheets_in_folder_or_workspace(self):
        """Returns the names and IDs of all sheets within a folder or workspace.
        
        Returns:
            list: A list of dictionaries, where each dictionary contains the name and ID of a sheet.
        """
        sheets = self.smartsheet_client.Folders.list_sheets(self.folder_or_workspace_id).result.data
        return [{"name": sheet.name, "id": sheet.id} for sheet in sheets]


    def get_sheet_id_by_name(self):
        """Returns the ID of a sheet by its name.
        
        Returns:
            int: The ID of the sheet.
        """
        sheets = self.get_sheets_in_folder_or_workspace()
        sheet = next((sheet for sheet in sheets if sheet["name"] == self.sheet_name), None)
        return sheet["id"] if sheet else None


    def create_sheet_from_template(self, template_sheet_id, new_sheet_name):
        """Creates a new sheet in a folder or workspace from a template.
        
        Args:
            template_sheet_id (int): The ID of the template sheet to use.
            new_sheet_name (str): The name of the new sheet.
        
        Returns:
            Sheet: The new sheet that was created.
        """
        new_sheet = self.smartsheet_client.Folders.create_sheet_from_template(
            self.folder_or_workspace_id,
            smartsheet.models.Sheet({
                'name': new_sheet_name,
                'from_id': int(template_sheet_id)
            })
        ).result

        return new_sheet


    def update_sheet_columns(self, sheet_id, new_columns):
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
