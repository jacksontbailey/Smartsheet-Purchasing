from core.config import settings

# - Let's create a helper function to create a smartsheet client
def create_smartsheet_client():
    # - First, let's build the credentials object
    from smartsheet import Smartsheet

    try:
        return Smartsheet(access_token=settings.API_KEY)
    except Exception as e:
        print(e)
        raise Exception("Error creating Smartsheet client")