import win32file
import win32api
import win32con
import wmi
import time
import os
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import base64
import json

class SecureUSBMonitor:
    def __init__(self, target_filename, secret_key, expected_content):
        """
        Initialize USB monitor with target filename and encryption parameters
        
        Args:
            target_filename (str): Name of the encrypted file to look for
            secret_key (str): Secret key for decryption
            expected_content (str): Content that should be in the file after decryption
        """
        self.target_filename = target_filename
        self.secret_key = secret_key
        self.expected_content = expected_content
        self.wmi = wmi.WMI()
        self.known_drives = set()
        
        # Initialize encryption key
        self.fernet = self.setup_encryption()

    def setup_encryption(self):
        """
        Setup encryption using the provided secret key
        """
        # Use PBKDF2 to derive a key from the secret
        salt = b'anthropic_salt_value_secure'  # In production, use a secure random salt
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(self.secret_key.encode()))
        return Fernet(key)

    def encrypt_content(self, content):
        """
        Encrypt content using Fernet (AES)
        """
        return self.fernet.encrypt(content.encode())

    def decrypt_content(self, encrypted_content):
        """
        Decrypt content using Fernet (AES)
        """
        try:
            decrypted_content = self.fernet.decrypt(encrypted_content)
            return decrypted_content.decode()
        except Exception as e:
            print(f"Decryption error: {e}")
            return None

    def get_usb_drives(self):
        """
        Get all connected USB drives
        """
        usb_drives = set()
        for drive in self.wmi.Win32_DiskDrive():
            if 'USB' in drive.InterfaceType:
                for partition in drive.associators("Win32_DiskDriveToDiskPartition"):
                    for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                        if logical_disk.DriveType == 2:  # Type 2 is removable drive
                            usb_drives.add(logical_disk.DeviceID)
        return usb_drives

    def verify_file_content(self, file_path):
        """
        Verify the content of the encrypted file
        """
        try:
            with open(file_path, 'rb') as file:
                encrypted_content = file.read()
                decrypted_content = self.decrypt_content(encrypted_content)
                
                if decrypted_content is None:
                    return False, "Decryption failed"
                
                # Try to parse JSON if the content is in JSON format
                try:
                    decrypted_content = json.loads(decrypted_content)
                except json.JSONDecodeError:
                    pass  # Content is not JSON, use as is
                
                return decrypted_content == self.expected_content, decrypted_content
        except Exception as e:
            print(f"Error reading file: {e}")
            return False, str(e)

    def search_and_verify_file(self, drive_letter):
        """
        Search for the target file and verify its contents
        """
        try:
            for root, _, files in os.walk(drive_letter + "\\"):
                if self.target_filename in files:
                    file_path = os.path.join(root, self.target_filename)
                    is_valid, content = self.verify_file_content(file_path)
                    return is_valid, content, file_path
        except Exception as e:
            print(f"Error searching in {drive_letter}: {e}")
        return False, None, None

    def get_drive_info(self, drive_letter):
        """
        Get drive information
        """
        try:
            volume_name = win32api.GetVolumeInformation(drive_letter + "\\")[0]
            sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters = \
                win32file.GetDiskFreeSpace(drive_letter + "\\")
            total_space = total_clusters * sectors_per_cluster * bytes_per_sector
            free_space = free_clusters * sectors_per_cluster * bytes_per_sector
            return volume_name, total_space, free_space
        except:
            return "Unknown", 0, 0

    def start_monitoring(self):
        """
        Start continuous monitoring of USB devices
        """
        print(f"Starting secure USB monitoring... Looking for file: {self.target_filename}")
        print("Waiting for USB drives...")

        while True:
            try:
                current_drives = self.get_usb_drives()
                
                # Check for new drives
                new_drives = current_drives - self.known_drives
                if new_drives:
                    for drive in new_drives:
                        print(f"\nNew USB drive detected: {drive}")
                        
                        # Get and display drive information
                        volume_name, total_space, free_space = self.get_drive_info(drive)
                        print(f"Volume Name: {volume_name}")
                        print(f"Total Space: {total_space / (1024**3):.2f} GB")
                        print(f"Free Space: {free_space / (1024**3):.2f} GB")
                        
                        # Search for and verify the file
                        print(f"Searching for and verifying {self.target_filename}...")
                        is_valid, content, file_path = self.search_and_verify_file(drive)
                        
                        if file_path:
                            if is_valid:
                                print("\n=== SUCCESS ===")
                                print(f"Valid encrypted file found at: {file_path}")
                                print("Content verification successful!")
                            else:
                                print("\n=== INVALID USB ===")
                                print(f"File found but content verification failed!")
                                print(f"Found at: {file_path}")
                        else:
                            print("\n=== INVALID USB ===")
                            print(f"Target file not found in {drive}")

                # Check for removed drives
                removed_drives = self.known_drives - current_drives
                if removed_drives:
                    for drive in removed_drives:
                        print(f"\nUSB drive removed: {drive}")

                # Update known drives
                self.known_drives = current_drives
                
                time.sleep(1)
                
            except Exception as e:
                print(f"Error in monitoring loop: {e}")
                time.sleep(1)

def create_encrypted_file(filename, content, secret_key):
    """
    Utility function to create an encrypted file for testing
    """
    monitor = SecureUSBMonitor(filename, secret_key, content)
    encrypted_content = monitor.encrypt_content(json.dumps(content))
    
    with open(filename, 'wb') as file:
        file.write(encrypted_content)
    print(f"Created encrypted file: {filename}")

def main():
    # Configuration
    target_file = "secure.enc"
    secret_key = "6RbAoW5b9U2GrKCoELi354XJ5lBPxwg7"  # Change this to your secure key
    expected_content = "valid-usb-token-2024"  # Change this to your expected content
    
    # Uncomment the following line to create a test encrypted file
    # create_encrypted_file(target_file, expected_content, secret_key)
    
    # Start monitoring
    usb_monitor = SecureUSBMonitor(target_file, secret_key, expected_content)
    usb_monitor.start_monitoring()

if __name__ == "__main__":
    main()