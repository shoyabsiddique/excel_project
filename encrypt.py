from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import base64

def generate_encrypted_file(output_filename, secret_key, content):
    """
    Generate an encrypted file using the same encryption method as the USB monitor
    
    Args:
        output_filename (str): Name of the file to create
        secret_key (str): Secret key for encryption
        content (str): Content to encrypt
    """
    # Setup encryption using PBKDF2
    salt = b'anthropic_salt_value_secure'
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    
    # Generate the key
    key = base64.urlsafe_b64encode(kdf.derive(secret_key.encode()))
    fernet = Fernet(key)
    
    # Encrypt the content
    encrypted_content = fernet.encrypt(content.encode())
    
    # Write to file
    with open(output_filename, 'wb') as file:
        file.write(encrypted_content)
    
    print(f"Successfully created encrypted file: {output_filename}")
    print(f"You can now copy this file to your USB drive")

if __name__ == "__main__":
    # Configuration - make sure these match your USB monitor script
    output_file = "secure.enc"
    secret_key = "6RbAoW5b9U2GrKCoELi354XJ5lBPxwg7"  # Change this to match your USB monitor script
    content_to_encrypt = "valid-usb-token-2024"  # Change this to match your USB monitor script
    
    # Generate the file
    generate_encrypted_file(output_file, secret_key, content_to_encrypt)