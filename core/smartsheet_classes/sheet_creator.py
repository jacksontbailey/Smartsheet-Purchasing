import smartsheet
from smartsheet_functions.client import create_smartsheet_client


class SheetCreator:
    def __init__(self, template_sheet_id, new_sheet_name, destination_workspace_id):
        self.template_sheet_id = template_sheet_id
        self.new_sheet_name = new_sheet_name
        self.destination_workspace_id = destination_workspace_id
        self.smartsheet_client = create_smartsheet_client()

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