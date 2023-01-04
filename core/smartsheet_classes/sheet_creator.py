import smartsheet
from smartsheet_functions.client import create_smartsheet_client


class SheetCreator:
    """Creates a new sheet in a Smartsheet workspace from a template.

    Args:
        api_key (str): The API key to use for authenticating the client.
        template_sheet_id (int): The ID of the template sheet to use.
        new_sheet_name (str): The name of the new sheet.
        destination_workspace_id (int): The ID of the destination workspace.

    Returns:
        Sheet: The new sheet that was created.
    """
    
    def __init__(self, api_key, template_sheet_id, new_sheet_name, destination_workspace_id):
        self.template_sheet_id = template_sheet_id
        self.new_sheet_name = new_sheet_name
        self.destination_workspace_id = destination_workspace_id
        self.smartsheet_client = create_smartsheet_client(api_key)

    def create_sheet(self):
        # - Create a new sheet from the template
        new_sheet = self.smartsheet_client.Workspaces.create_sheet_in_workspace_from_template(
            self.destination_workspace_id,
            smartsheet.models.Sheet({
                'name': self.new_sheet_name,
                'from_id': int(self.template_sheet_id)
            })
        ).result

        # - Return the new sheet
        return new_sheet