import configparser
import os
import tkinter as tk
from tkinter import messagebox

# Function to create a session in the winscp.ini file
def create_session(host_name, transfer_type, folder_name, session_name):
    # Path to the winscp.ini file
    ini_path = os.path.expanduser('~\\AppData\\Roaming\\WinSCP.ini')

    # Check if the INI file exists
    if not os.path.exists(ini_path):
        return f"INI file not found at {ini_path}"

    # Create config parser and read the INI file
    config = configparser.ConfigParser()
    config.read(ini_path)

    # Define the section name based on the folder and session name
    section_name = f'Sessions\\"{folder_name}"/"{session_name}"'

    # Check if the session already exists
    if config.has_section(section_name):
        return "Session already exists."

    # Define session details
    config[section_name] = {
        'HostName': host_name,
        'UserName': 'your_username',
        'Password': 'your_encrypted_password',  # Replace with real encrypted password
        'PortNumber': '20022' if transfer_type.lower() == 'sftp' else '21',
    }

    # Add FSProtocol only for FTP
    if transfer_type.lower() == 'ftp':
        config[section_name]['FSProtocol'] = '5'

    # Write the session to the INI file
    with open(ini_path, 'w') as configfile:
        config.write(configfile)

    return "Session created successfully!"

# Function to handle button click in Tkinter
def on_create():
    host_name = host_entry.get().strip()
    transfer_type = type_var.get().strip()
    folder_name = folder_entry.get().strip()
    session_name = session_entry.get().strip()

    if not host_name or not transfer_type or not folder_name or not session_name:
        messagebox.showwarning("Input Error", "Please provide all required inputs.")
        return

    result = create_session(host_name, transfer_type, folder_name, session_name)
    
    if "successfully" in result:
        messagebox.showinfo("WinSCP Session", result)
    else:
        messagebox.showerror("Error", result)

# Create the main application window
root = tk.Tk()
root.title("WinSCP Session Creator")

# Create input fields and labels
tk.Label(root, text="Host Name:").pack(pady=5)
host_entry = tk.Entry(root, width=50)
host_entry.pack(pady=5)

tk.Label(root, text="Folder Name:").pack(pady=5)
folder_entry = tk.Entry(root, width=50)
folder_entry.pack(pady=5)

tk.Label(root, text="Session Name:").pack(pady=5)
session_entry = tk.Entry(root, width=50)
session_entry.pack(pady=5)

tk.Label(root, text="Transfer Type:").pack(pady=5)
type_var = tk.StringVar(value="sftp")
transfer_type_dropdown = tk.OptionMenu(root, type_var, "sftp", "ftp")
transfer_type_dropdown.pack(pady=5)

# Create a button to trigger the session creation
create_button = tk.Button(root, text="Create Session", command=on_create)
create_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
