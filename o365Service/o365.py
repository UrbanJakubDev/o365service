import requests
import os
from dotenv import load_dotenv
import msal
from dataclasses import dataclass
from typing import Optional
import os
import shutil
from pathlib import Path

# Load environment variables from .env file
load_dotenv()

# Add logger to the service
import logging

logger = logging.getLogger(__name__)



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
            logger.error(f"Folder '{sub_folder}' not found")

    def download_file(self, site_id, drive_id, file, download_path):
        file_id = file['id']
        file_name = file['name']
        logger.info(f'Downloading file: {file_name}')

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
            except Exception as e:
                logger.error(f'Failed to download file: {str(e)}')
        else:
            logger.error(f'Failed to download file: {response.text}')




@dataclass
class SharepointConfig:
    """Configuration for Sharepoint service"""
    tenant_id: str
    client_id: str
    client_secret: str
    download_path: str
    site_id: str = "401afd8d-64ec-4e9d-87ec-1a39fd2a4c58"
    drive_id: str = "b!jf0aQOxknU6H7Bo5_SpMWMjp5nA7vjNKtxvdrybzT3dEaDML2YwjS7LwfrvufMoA"
    resource_url: str = 'https://graph.microsoft.com/v1.0/sites'
    folder_prefix: str = 'Provozní hodnoty '
    subfolder: str = "Jednotky"

class SharepointService:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        download_path: str,
        config: Optional[SharepointConfig] = None
    ) -> None:
        """
        Initialize the SharepointService with required parameters.

        Args:
            tenant_id (str): The ID of the tenant
            client_id (str): The ID of the client
            client_secret (str): The secret key of the client
            download_path (str): Path where files will be downloaded
            config (Optional[SharepointConfig]): Optional configuration object

        Raises:
            ValueError: If any required parameter is missing or invalid
        """
        # Validate required parameters
        if not all([tenant_id, client_id, client_secret, download_path]):
            raise ValueError(
                "All parameters (tenant_id, client_id, client_secret, download_path) "
                "are required"
            )

        if config:
            self.config = config
        else:
            self.config = SharepointConfig(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                download_path=download_path
            )

        # Set instance attributes
        self.tenant_id = self.config.tenant_id
        self.client_id = self.config.client_id
        self.client_secret = self.config.client_secret
        self.site_id = self.config.site_id
        self.drive_id = self.config.drive_id
        self.resource_url = self.config.resource_url
        self.download_path = Path(self.config.download_path).resolve()
        self.folder = self.config.folder_prefix
        self.subfolder = self.config.subfolder

        # Validate download path
        if not self._validate_download_path():
            raise ValueError(
                f"Download path '{self.download_path}' is not valid or not accessible"
            )

    def _validate_download_path(self) -> bool:
        """
        Validate that the download path is valid and accessible.
        
        Returns:
            bool: True if path is valid and accessible
        """
        try:
            self.download_path.mkdir(parents=True, exist_ok=True)
            return self.download_path.is_dir() and os.access(str(self.download_path), os.W_OK)
        except Exception:
            return False

    def run(self, year: int) -> None:
        """
        Run the Sharepoint service to download files.

        Args:
            year (int): The year for which to download files

        Raises:
            ValueError: If year is not valid
            Exception: If any error occurs during execution
        """
        if not isinstance(year, int) or year <= 0:
            raise ValueError(f"Invalid year value: {year}")

        try:
            # Create SharePointClient instance
            sharepoint_client = SharePointClient(
                self.tenant_id,
                self.client_id,
                self.client_secret,
                self.resource_url,
                year
            )

            # Get access token
            sharepoint_client.get_access_token()

            # Get drive contents
            sharepoint_client.drive_folders = sharepoint_client.get_drive_content(
                self.site_id,
                self.drive_id
            )

            # Get folder contents
            folder_name = f"{self.folder}{year}"
            files_folder_id = sharepoint_client.get_folder_id_by_name(folder_name)
            if not files_folder_id:
                raise ValueError(f"Folder not found: {folder_name}")

            files_folder_contents = sharepoint_client.get_folder_content(
                self.site_id,
                self.drive_id,
                files_folder_id
            )

            # Get files from subfolder
            sharepoint_client.get_files_from_folder(
                self.site_id,
                self.drive_id,
                files_folder_contents,
                self.subfolder
            )

            # Download files
            for file in sharepoint_client.files:
                sharepoint_client.download_file(
                    self.site_id,
                    self.drive_id,
                    file,
                    str(self.download_path)
                )

        except Exception as e:
            raise Exception(f"Error during Sharepoint operation: {str(e)}")

    def dropDataFolder(self) -> None:
        """
        Remove the download folder and all its contents.
        
        Raises:
            Exception: If an error occurs while removing the folder
        """
        try:
            if self.download_path.exists():
                shutil.rmtree(str(self.download_path))
        except Exception as e:
            raise Exception(f"Error while removing data folder: {str(e)}")