import configparser
import os
import tkinter as tk
from tkinter import messagebox

# Function to create a session in the winscp.ini file
def create_session(host_name, transfer_type, folder_name, session_name):
    # Dynamic path based on the folder name
    custom_ini_dir = os.path.join(r'C:\WinSCPConfig', folder_name)
    custom_ini_path = os.path.join(custom_ini_dir, 'WinSCP.ini')

    # Ensure the directory exists
    os.makedirs(custom_ini_dir, exist_ok=True)

    # Create config parser and read the INI file (if it exists)
    config = configparser.ConfigParser()
    if os.path.exists(custom_ini_path):
        config.read(custom_ini_path)

    # Define the section name based on the folder and session name
    section_name = f'Sessions\\{folder_name}/{session_name}'

    # Check if the session already exists
    if config.has_section(section_name):
        return "Session already exists."

    # Define session details
    config[section_name] = {
        'HostName': host_name,
        'PortNumber': '20022' if transfer_type.lower() == 'sftp' else '21',
        'UserName': 'Administrator' if transfer_type.lower() == 'sftp' else 'anonymous',
        'Password': 'A35C45504648113EE96A1003AC13A5A41D38313532352F282E3D28332E6D6B6E726E6C726E726A6A6D84CA5BFA50425E8C85' if transfer_type.lower() == 'sftp' else 'A35C755E6D593D323332253133292F6D6B6E726E6C726E72696D3D323332253133292F1C39243D312C3039723F333130FAB0',
    }

    # Add FSProtocol only for FTP
    if transfer_type.lower() == 'ftp':
        config[section_name]['FSProtocol'] = '5'

    # Write the session to the INI file
    with open(custom_ini_path, 'w') as configfile:
        config.write(configfile)

    return f"Session created successfully in {custom_ini_path}!"

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

# Function to open the directory containing the INI file
def open_ini_directory():
    folder_name = folder_entry.get().strip()
    custom_ini_dir = os.path.join(r'C:\WinSCPConfig', folder_name)
    
    if os.path.exists(custom_ini_dir):
        os.startfile(custom_ini_dir)
    else:
        messagebox.showerror("Error", f"Directory not found: {custom_ini_dir}")

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
create_button.pack(pady=10)

# Create a button to open the directory containing the INI file
open_dir_button = tk.Button(root, text="Open INI Directory", command=open_ini_directory)
open_dir_button.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
