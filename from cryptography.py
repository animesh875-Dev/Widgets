# from cryptography.fernet import Fernet

# # Load the existing key (do not regenerate it)
# with open("key.key", "rb") as key_file:
#     key = key_file.read()

# cipher_suite = Fernet(key)
# # Encrypt the password
# password = "Password@875"
# encrypted_password = cipher_suite.encrypt(password.encode())
# print(f"Encrypted Password: {encrypted_password.decode()}")

from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
import base64

# Key and data must be of appropriate sizes
key = b"thisisaverysecure"  # 16 bytes for AES-128
password = "Password@875"

# Encrypt
cipher = AES.new(key, AES.MODE_CBC)
ct_bytes = cipher.encrypt(pad(password.encode(), AES.block_size))
encrypted_password = base64.b64encode(cipher.iv + ct_bytes).decode()
print(f"Encrypted Password: {encrypted_password}")

# Decrypt
data = base64.b64decode(encrypted_password)
iv = data[:AES.block_size]
ct = data[AES.block_size:]
cipher = AES.new(key, AES.MODE_CBC, iv)
decrypted_password = unpad(cipher.decrypt(ct), AES.block_size).decode()
print(f"Decrypted Password: {decrypted_password}")
