from smartsheet import Smartsheet

def create_smartsheet_client(api_key):
    """Creates a new Smartsheet client.

    Args:
        api_key (str): The API key to use for authenticating the client.

    Returns:
        SmartsheetClient: The authenticated Smartsheet client.
    """
    # - Set up the Smartsheet client
    ss_client = Smartsheet(access_token = api_key)
    ss_client.errors_as_exceptions(True)

    return ss_client