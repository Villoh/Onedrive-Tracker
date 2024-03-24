# OneDrive funcs
import os
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.runtime.client_request_exception import ClientRequestException

def download_files(remote_folder: DriveItem, local_path: str) -> None:
    """Download files from a given DriveItem (Object Path in OneDrive/Sharepoint)

    Args:
        remote_folder (DriveItem): Starting path where it's going to start looking for files.
        local_path (str): Path where files will be stored.
    """
    try:
        drive_items = remote_folder.children.get().execute_query()
        for drive_item in drive_items:
            if not drive_item.is_file:
                print(f"Searching through folder: {drive_item.name} from {drive_item.web_url}")
                download_files(drive_item, f'{local_path}/{drive_item.name}/')
            else:
                print(f"Downloading file: {drive_item.name} into: {local_path} from {drive_item.web_url}")
                # download file content
                os.makedirs(name=local_path, exist_ok=True)
                with open(os.path.join(local_path, drive_item.name), 'wb') as local_file:
                    drive_item.download(local_file).execute_query()
    except ClientRequestException as e:
        print(f"Error: {e}")