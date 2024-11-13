import os
import platform
import sys

def is_colab():
    """Check if code is running in Google Colab"""
    try:
        import google.colab
        return True
    except ImportError:
        return False

def is_gdrive_path(path):
    """Check if the path is a Google Drive path"""
    # Common Google Drive mount points
    gdrive_indicators = [
        '/content',
        '/content/drive',
        '/content/shared_drive',
        'drive/My Drive',
        'drive/MyDrive'
    ]
    return any(indicator in path for indicator in gdrive_indicators)

def safe_move(path1, path2):
    """
    Cross-platform safe move operation
    
    Args:
        path1 (str): Source path
        path2 (str): Destination path
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Check if we're in Colab and dealing with Google Drive paths
        if is_colab() and (is_gdrive_path(path1) or is_gdrive_path(path2)):
            print(f"[debug] system: google colab")
            try:
                from gdrive_utils import safe_move_file  # Your Google Drive specific code
                return safe_move_file(path1, path2)
            except ImportError:
                print("Google Drive utilities not found. Please ensure gdrive_utils.py is available.")
                return False
        
        # For local file system operations
        else:
            # Handle different operating systems
            system = platform.system().lower()
            print(f"[debug] system: {system}")
            
            if system == 'darwin':  # macOS
                os.rename(path1, path2)
            elif system == 'windows':
                # Windows might need special handling for certain paths
                os.replace(path1, path2)  # More robust than rename on Windows
            else:  # Linux and others
                os.rename(path1, path2)
            
            return True
            
    except Exception as e:
        print(f"Error moving file: {str(e)}")
        return False

def move_file(path1, path2):
    """
    User-friendly wrapper for file moving operations
    """
    if safe_move(path1, path2):
        print(f"- moved file from {path1} to {path2}")
    else:
        print("Failed to move file")
