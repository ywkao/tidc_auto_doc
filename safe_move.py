from google.colab import drive
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os

class SharedDriveMover:
    def __init__(self):
        # Add shared drive scope
        self.SCOPES = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/drive.metadata',
        ]
        self.service = self._authenticate()
        
    def _authenticate(self):
        """Set up Google Drive API service"""
        if not os.path.exists('/content/drive'):
            drive.mount('/content/drive')
            
        # Your authentication code here
        return build('drive', 'v3', credentials=self.creds)
    
    def _get_shared_drive_id(self, drive_name):
        """Get the ID of the shared drive"""
        results = self.service.drives().list(
            pageSize=50
        ).execute()
        
        for drive_item in results.get('drives', []):
            if drive_name in drive_item['name']:
                return drive_item['id']
        return None
    
    def _get_folder_id(self, folder_path, shared_drive_id):
        """Get folder ID from path in shared drive"""
        clean_path = folder_path.replace('/content/drive/', '')
        path_parts = clean_path.strip('/').split('/')
        
        # Start from shared drive root
        parent_id = shared_drive_id
        
        for part in path_parts:
            query = f"name='{part}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents"
            results = self.service.files().list(
                q=query,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                corpora='drive',
                driveId=shared_drive_id,
                fields='files(id, name)',
                spaces='drive'
            ).execute()
            
            items = results.get('files', [])
            if not items:
                raise Exception(f"Could not find folder: {part}")
            
            parent_id = items[0]['id']
        
        return parent_id
    
    def move_file(self, file_name, source_path, target_path, shared_drive_name):
        """
        Move a file between folders in a shared drive
        
        Args:
            file_name (str): Name of the file to move
            source_path (str): Full path to source folder
            target_path (str): Full path to target folder
            shared_drive_name (str): Name of the shared drive
        """
        try:
            # Get shared drive ID
            shared_drive_id = self._get_shared_drive_id(shared_drive_name)
            if not shared_drive_id:
                raise Exception(f"Could not find shared drive: {shared_drive_name}")
            
            # Get source and target folder IDs
            source_folder_id = self._get_folder_id(source_path, shared_drive_id)
            target_folder_id = self._get_folder_id(target_path, shared_drive_id)
            
            # Find the file in the source folder
            query = f"name='{file_name}' and '{source_folder_id}' in parents"
            results = self.service.files().list(
                q=query,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                corpora='drive',
                driveId=shared_drive_id,
                fields='files(id, name)'
            ).execute()
            
            items = results.get('files', [])
            if not items:
                raise Exception(f"Could not find file: {file_name}")
            
            file_id = items[0]['id']
            
            # Move the file
            self.service.files().update(
                fileId=file_id,
                addParents=target_folder_id,
                removeParents=source_folder_id,
                supportsAllDrives=True
            ).execute()
            
            print(f"Successfully moved {file_name} to {target_path}")
            return True
            
        except Exception as e:
            print(f"Error moving file: {str(e)}")
            return False

# Usage example
def move_shared_drive_file(shared_drive_name, source, target):
    dir_source = os.path.dirname(source)
    dir_target = os.path.dirname(target)
    file_name = os.path.basename(source)

    mover = SharedDriveMover()
    return mover.move_file(
        file_name=file_name,
        source_path=dir_source,
        target_path=dir_target,
        shared_drive_name=shared_drive_name
    )
