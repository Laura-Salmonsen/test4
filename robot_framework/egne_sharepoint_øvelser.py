import pandas as pd
import os
from openpyxl import Workbook
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
def create_excel_on_sharepoint(client: ClientContext, sharepoint_folder_url: str, file_name: str, sheet_data = None) -> str:
    """
    Creates a new Excel file with optional data and uploads it to a specified SharePoint folder.
    
    Parameters:
    - client: ClientContext object for SharePoint connection.
    - sharepoint_folder_url: The SharePoint folder URL where the file will be created.
    - file_name: Name of the new Excel file (e.g., 'new_file.xlsx').
    - sheet_data: Optional list of lists representing rows of data to populate the Excel sheet. Each inner list is a row.

    Returns:
    - URL of the uploaded file on SharePoint.
    """
    if not file_name.endswith(".xlxs"):
        file_name += ".xlsx"
    
    local_file_path = os.path.join(os.getcwd(), file_name)
    workbook = Workbook()
    sheet = workbook.active

    if sheet_data:
        for row in sheet_data:
            sheet.append(row)
    
    workbook.save(local_file_path)

    #An excel file has now been created - either empty or optionally containing data. 

    try:
        upload_file_to_sharepoint(client = client, local_file_path = local_file_path, sharepoint_folder_url= sharepoint_folder_url, new_file_name= None )
    
    finally:
        if os.path.exists(local_file_path):
            os.remove(local_file_path)
def sharepoint_client(username: str, password: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    orchestrator_connection.log_info(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx
def download_file_from_sharepoint(client: ClientContext, sharepoint_file_url: str) -> str:
    """
    Downloads a file from SharePoint and returns the local file path.
    """
    file_name = sharepoint_file_url.split("/")[-1]
    download_path = os.path.join(os.getcwd(), file_name)

    # Download the file
    with open(download_path, "wb") as local_file:
        file = client.web.get_file_by_server_relative_path(sharepoint_file_url).download(local_file).execute_query()
    
    print(f"[Ok] file has been downloaded into: {download_path}")
    return download_path
def upload_file_to_sharepoint(client: ClientContext, local_file_path: str, sharepoint_folder_url: str, new_file_name: str) -> str:
    """
    Uploads file to sharepoint and returns URL of uploaded file
    """

    original_file_name = os.path.basename(local_file_path)
    file_name = new_file_name if new_file_name else original_file_name
    target_url = f"{sharepoint_folder_url}/{file_name}"

    # The upload
    with open(local_file_path, "rb") as local_file:
        target_folder = client.web.get_folder_by_server_relative_url(sharepoint_folder_url)
        target_folder.upload_file(file_name, local_file.read()).execute_query()

    print(f"[Ok] file has been uploaded to: {target_url}")
    return target_url
def move_file_in_sharepoint(client: ClientContext, from_url: str, to_url: str) -> str:
    file_name = from_url.split('/')[-1]
    target_file_url = f"{to_url}/{file_name}"

    file = client.web.get_file_by_server_relative_url(from_url)
    file.moveto(to_url, 1).execute_query()

    print(f"Filen er flyttelyttet til {to_url}")

def copy_file_in_sharepoint(client: ClientContext, from_url: str, to_folder_url: str) -> str:
    """
    Copies a file from one location to another in SharePoint and adds '_copy' to the filename.
    Raises an error if a file with the same name already exists in the destination folder.

    Parameters:
    - client: ClientContext object for SharePoint connection.
    - from_url: The server-relative URL of the file to copy.
    - to_folder_url: The server-relative URL of the destination folder.

    Returns:
    - The server-relative URL of the copied file in SharePoint.
    """
    # Extract the file name from the source URL
    file_name = from_url.split('/')[-1]

    # Split the file name into base name and extension
    base_name, extension = os.path.splitext(file_name)

    # Add '_copy' to the base name
    new_file_name = f"{base_name}_copy{extension}"

    # Construct the full destination file URL
    target_file_url = f"{to_folder_url}/{new_file_name}"

    # Check if a file with the same name already exists in the destination folder
    try:
        client.web.get_file_by_server_relative_url(target_file_url).get().execute_query()
        raise FileExistsError(f"A file named '{new_file_name}' already exists in the folder '{to_folder_url}'.")
    except Exception as e:
        # If the exception is not a 404 (file not found), re-raise the exception
        if "404" not in str(e):
            raise

    # Perform the copy operation
    file = client.web.get_file_by_server_relative_url(from_url)
    file.copyto(target_file_url).execute_query()

    print(f"Filen er kopireliret {target_file_url}")
    return target_file_url


# Example usage:

# Get credentials from Orchestrator
orchestrator_connection = OrchestratorConnection("Test-sharepoint-sjov", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)
RobotCredentials = orchestrator_connection.get_credential("RobotCredentials")
username = RobotCredentials.username
password = RobotCredentials.password

# SharePoint site and parent folder URL
SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com/Teams/tea-teamsite11819"
PARENT_FOLDER_URL = "/Teams/tea-teamsite11819/Delte Dokumenter/Testmappe/testmappe3"
to_url = "/Teams/tea-teamsite11819/Delte Dokumenter/Testmappe/testmappe2"
name = "newly_created_file.xlsx"

# Create the SharePoint client
client = sharepoint_client(username, password, SHAREPOINT_SITE_URL, orchestrator_connection = orchestrator_connection)
local_file_path = "C:/Users/az81549/Downloads/GEO-2034-234566.xlsx"

# result = copy_file_in_sharepoint(client = client, from_url= PARENT_FOLDER_URL +'/' + name, to_folder_url= to_url )

print(username, password)
