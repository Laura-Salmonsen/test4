import os
import pandas as pd
import re
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext




# def lav_ny_mappe(client: ClientContext, parent_folder_url = str, mappenavn: str,  orchestrator_connection: OrchestratorConnection)
    
def tjek_for_aktindsigt(client: ClientContext, parent_folder_url: str, orchestrator_connection: OrchestratorConnection):
    """
    Traverses folders under the specified parent folder in SharePoint, checks each Excel file for a specific column,
    and returns a dictionary with folder paths and their check results.
    """
    orchestrator_connection.log_trace("Running tjek_for_aktindsigt.")
    
    # Dictionary to store results for each folder
    results = {}

    # Recursively traverse folders and check Excel files
    traverse_and_check_folders(client, parent_folder_url, results, orchestrator_connection)

    return results

def traverse_and_check_folders(client, folder_url, results, orchestrator_connection):
    """
    Recursively traverses through folders in SharePoint, filters by matching folder names,
    checks for Excel files, and saves check results in `results`.
    """
    # Define the regex pattern for folder names (e.g., "GEO-2024-123456")
    pattern = re.compile(r"^[A-Z]{3}-\d{4}-\d{6}$")

    folder = client.web.get_folder_by_server_relative_url(folder_url)
    client.load(folder)
    client.execute_query()

    # Load subfolders in the current folder
    subfolders = folder.folders
    client.load(subfolders)
    client.execute_query()

    # Process each subfolder if it matches the pattern
    for subfolder in subfolders:
        subfolder_name = subfolder.properties["Name"]
        subfolder_url = f"{folder_url}/{subfolder_name}"

        # Only proceed if the folder name matches the specified pattern
        if pattern.match(subfolder_name):
            # orchestrator_connection.log_info(f"Checking folder: {subfolder_url}") - orker ikke alle de logs
            # Check each file in the matched folder
            files = subfolder.files
            client.load(files)
            client.execute_query()

            for file in files:
                if file.properties["Name"].endswith(".xlsx"):
                    # Download and process the Excel file
                    file_url = f"{subfolder_url}/{file.properties['Name']}"
                    local_file_path = download_file_from_sharepoint(client, file_url)
                    result = check_excel_file(local_file_path, orchestrator_connection)
                    
                    # Store result using only the folder name as the key
                    results[subfolder_name] = result
                    
                    os.remove(local_file_path)  # Clean up after processing
                    break  # Stop after processing the first Excel file in this folder

        # Recursively traverse the subfolders to continue searching deeper
        traverse_and_check_folders(client, subfolder_url, results, orchestrator_connection)



def check_excel_file(file_path: str, orchestrator_connection: OrchestratorConnection) -> str:
    """
    Checks the 'Gives der aktindsigt?' column in the specified Excel file and returns the result.
    """
    try:
        df = pd.read_excel(file_path)
        

        # Strip any leading or trailing whitespace from column names
        df.columns = df.columns.str.strip()

        if 'Gives der aktindsigt?' in df.columns:
            if (df["Gives der aktindsigt?"] == "Ja").all():
                return 'Fuld aktindsigt'
            elif (df["Gives der aktindsigt?"] == "Nej").all():
                return 'Afvist'
            else:
                return 'Delvis aktindsigt'
        else:
            orchestrator_connection.log_error("Column 'Gives der aktindsigt?' not found in the file.")
            return "Column 'Gives der aktindsigt?' not found"
    except Exception as e:
        orchestrator_connection.log_error(f"Error reading Excel file at {file_path}: {e}")
        return "Error processing file"

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

# Example usage:

# Get credentials from Orchestrator
orchestrator_connection = OrchestratorConnection("Test_Laura", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)
RobotCredentials = orchestrator_connection.get_credential("RobotCredentials")
username = RobotCredentials.username
password = RobotCredentials.password

# SharePoint site and parent folder URL
SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com/Teams/tea-teamsite11819"
PARENT_FOLDER_URL = "/Teams/tea-teamsite11819/Delte Dokumenter/Testmappe/testmappe2"

# Create the SharePoint client
client = sharepoint_client(username, password, SHAREPOINT_SITE_URL, orchestrator_connection = orchestrator_connection)

# Run tjek_for_aktindsigt to check each Excel file in subfolders

results = tjek_for_aktindsigt(client, PARENT_FOLDER_URL, orchestrator_connection)
print("Results:", results)



