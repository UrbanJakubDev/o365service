import requests
import os
from dotenv import load_dotenv
import msal

# Load environment variables from .env file
load_dotenv()


class SharePointClient:
    def __init__(self, tenant_id, client_id, client_secret, resource_url, year):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.resource_url = resource_url
        self.scope = ['https://graph.microsoft.com/.default']
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.base_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        self.headers = {'Content-Type': 'application/x-www-form-urlencoded'}

        self.client = msal.ConfidentialClientApplication(
            client_id, authority=self.authority, client_credential=client_secret)
        self.access_token = None
        self.drive_folders = []
        self.files = []
        self.year_folder = year

    def get_access_token(self):

        token_result = self.client.acquire_token_silent(
            self.scope, account=None)

        if token_result:
            self.access_token = token_result['access_token']

        if not token_result:
            token_result = self.client.acquire_token_for_client(
                scopes=self.scope)
            self.access_token = token_result['access_token']

    def get_drive_content(self, site_id, drive_id):
        # Get the contents of a drive
        drive_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children/General/children'

        response = requests.get(
            drive_url, headers={'Authorization': f'Bearer {self.access_token}'})
        items_data = response.json()

        items = []
        if 'value' in items_data:
            for item in items_data['value']:
                item_info = {
                    'name': item['name'],
                    'url': item['webUrl'],
                    'id': item['id'],
                    'folder': 'folder' in item  # Check if the item is a folder
                }
                items.append(item_info)
        return items

    def get_folder_id_by_name(self, folder_name):
        return next((item['id'] for item in self.drive_folders if item['name'] == folder_name), None)

    def get_folder_content(self, site_id, drive_id, folder_id='root'):
        # Get the contents of a folder
        folder_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children'

        response = requests.get(folder_url, headers={
                                'Authorization': f'Bearer {self.access_token}'})
        items_data = response.json()

        items = []
        if 'value' in items_data:
            for item in items_data['value']:
                item_info = {
                    'name': item['name'],
                    'url': item['webUrl'],
                    'id': item['id'],
                    'folder': 'folder' in item  # Check if the item is a folder
                }
                items.append(item_info)
        return items

    def get_files_from_folder(self, site_id, drive_id, main_folder=None, sub_folder=None):
        # Find the "Jednotky" folder within "Provozní hodnoty 2024"
        jednotky_folder_id = None
        for item in main_folder:
            if item['name'] == sub_folder and item['folder']:
                jednotky_folder_id = item['id']
                break

        if jednotky_folder_id:
            # Get the contents of the "Jednotky" folder
            self.files = self.get_folder_content(
                site_id, drive_id, jednotky_folder_id)
        else:
            print("Jednotky folder not found.")

    def download_file(self, site_id, drive_id, file, download_path):
        file_id = file['id']
        file_name = file['name']
        print(f'Processing file: {file_name}')

        # Skip if the file is a folder or a shortcut or doesn't have "Provozní hodnoty 2024" in the path or is not a .xlsx file
        if (not file_name.endswith('.xlsx') and not file_name.endswith('.xlsm')):
            return

        # Ensure the download path exists
        os.makedirs(download_path, exist_ok=True)

        # Download the file
        download_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}/content'
        response = requests.get(download_url, headers={
                                'Authorization': f'Bearer {self.access_token}'}, stream=True)

        if response.status_code == 200:
            try:
                with open(os.path.join(download_path, file_name), 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                print(f'File downloaded successfully to {download_path}')
            except Exception as e:
                print(f'Error writing file {file_name}: {e}')
        else:
            print(f'Failed to download file: {response.status_code}')


class SharepointService:

    def __init__(self, *args, **kwargs):
        """
        Initializes an instance of the O365 class.
        Args:
            tenant_id (str): The ID of the tenant.
            client_id (str): The ID of the client.
            client_secret (str): The secret key of the client.
            year (str, optional): The year for the folder name. Defaults to None.
        """

        # Get the value from args, kwargs
        self.tenant_id = kwargs["tenant_id"]
        self.client_id = kwargs["client_id"]
        self.client_secret = kwargs["client_secret"]
        self.site_id = "401afd8d-64ec-4e9d-87ec-1a39fd2a4c58"
        self.drive_id = "b!jf0aQOxknU6H7Bo5_SpMWMjp5nA7vjNKtxvdrybzT3dEaDML2YwjS7LwfrvufMoA"
        self.resource_url = 'https://graph.microsoft.com/v1.0/sites'
        self.download_path = './dataFiles/'
        self.folder = f'Provozní hodnoty '
        self.subfolder = "Jednotky"

    def run(self, year: int = 0):
        # Usage of the class
        site_id = self.site_id
        drive_id = self.drive_id

        # Create an instance of the SharePointClient class
        sharepoint_client = SharePointClient(self.tenant_id, self.client_id, self.client_secret, self.resource_url, year)
        sharepoint_client.get_access_token()

        # Get the contents of the root folder of the drive
        sharepoint_client.drive_folders = sharepoint_client.get_drive_content(site_id, drive_id)

        # Get the contents of the folder "Provozní hodnoty 2024"
        files_folder_id = sharepoint_client.get_folder_id_by_name(self.folder + str(year))
        files_folder_contents = sharepoint_client.get_folder_content(site_id, drive_id, files_folder_id)

        # Get the contents of the "Jednotky" folder
        sharepoint_client.get_files_from_folder(site_id, drive_id, files_folder_contents, self.subfolder)

        # Make sure the download path exists
        os.makedirs(self.download_path, exist_ok=True)

        # Download the files
        for file in sharepoint_client.files:
            sharepoint_client.download_file(site_id, drive_id, file, self.download_path)
            
    def dropDataFolder(self):
        import shutil
        shutil.rmtree(self.download_path)
