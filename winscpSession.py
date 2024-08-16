import configparser
import os
import winreg as reg
import tkinter as tk
from tkinter import messagebox
import subprocess
import psutil

# Function to set the custom INI path in the Windows Registry
def set_custom_ini_path(ini_path):
    key_path = r'Software\Martin Prikryl\WinSCP 2\Configuration'
    
    try:
        key = reg.OpenKey(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_SET_VALUE)
        reg.SetValueEx(key, "ConfigurationStorage", 0, reg.REG_DWORD, 2)
        reg.SetValueEx(key, "CustomIniFile", 0, reg.REG_SZ, ini_path)
        reg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Failed to set registry key: {e}")
        return False

# Function to close WinSCP if it's running
def close_winscp():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == "WinSCP.exe":
            proc.terminate()
            proc.wait()

# Function to create a session in the winscp.ini file
def create_session(host_name, transfer_type, folder_name, session_name):
    custom_ini_dir = os.path.join(r'C:\WinSCPConfig', folder_name)
    custom_ini_path = os.path.join(custom_ini_dir, 'WinSCP.ini')

    os.makedirs(custom_ini_dir, exist_ok=True)

    if not set_custom_ini_path(custom_ini_path):
        return "Failed to set the custom INI path in the registry."

    close_winscp()  # Close WinSCP to ensure settings take effect

    config = configparser.ConfigParser()
    if os.path.exists(custom_ini_path):
        config.read(custom_ini_path)

    section_name = f'Sessions\\{folder_name}/{session_name}'

    if config.has_section(section_name):
        return "Session already exists."

    config[section_name] = {
        'HostName': host_name,
        'PortNumber': '20022' if transfer_type.lower() == 'sftp' else '21',
        'UserName': 'Administrator' if transfer_type.lower() == 'sftp' else 'anonymous',
        'Password': 'A35C45504648113EE96A1003AC13A5A41D38313532352F282E3D28332E6D6B6E726E6C726E726A6A6D84CA5BFA50425E8C85' if transfer_type.lower() == 'sftp' else 'A35C755E6D593D323332253133292F6D6B6E726E6C726E72696D3D323332253133292F1C39243D312C3039723F333130FAB0',
    }

    if transfer_type.lower() == 'ftp':
        config[section_name]['FSProtocol'] = '5'

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

create_button = tk.Button(root, text="Create Session", command=on_create)
create_button.pack(pady=10)

open_dir_button = tk.Button(root, text="Open INI Directory", command=open_ini_directory)
open_dir_button.pack(pady=10)

root.mainloop()
