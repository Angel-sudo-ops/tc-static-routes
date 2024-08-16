import os
import tkinter as tk
from tkinter import messagebox

import os

# Path to the WinSCP executable
winscp_path = r'"C:\Program Files (x86)\WinSCP\winscp.com"'

def create_winscp_session():
    # Get the selected protocol
    protocol = protocol_var.get()
    
    # Hardcoded session details
    host = 'example.com'
    port = '22' if protocol == 'sftp' else '21'
    username = 'Administrator'
    password = 'your_password'
    privatekey = 'C:\\path\\to\\private_key.ppk' if protocol == 'sftp' else ''
    hostkey = 'ssh-rsa 2048 ABCDEFGHIJKLMNOPQRSTUVWXYZ' if protocol == 'sftp' else ''
    remotedir = '/home/your_username'
    localdir = 'C:\\Users\\your_user\\Documents'

    # Create a WinSCP script
    script_content = f"""
    open {protocol}://{username}:{password}@{host}:{port}/"""
    
    if protocol == 'sftp':
        script_content += f""" -privatekey={privatekey} -hostkey="{hostkey}" """
    
    script_content += f"""
    lcd {localdir}
    cd {remotedir}
    exit
    """

    # Save the script to a file
    script_filename = 'winscp_script.txt'
    with open(script_filename, 'w') as script_file:
        script_file.write(script_content)

    # Execute the WinSCP script via command line to save the session
    # os.system(f'winscp.com /script={script_filename} /save')
    os.system(f'{winscp_path} /script={script_filename} /save')

    # Show a message box to indicate the session has been created
    messagebox.showinfo("WinSCP Session", f"{protocol.upper()} session has been successfully created and saved!")

# Create the main application window
root = tk.Tk()
root.title("WinSCP Session Creator")

# Create a label and dropdown for protocol selection
protocol_label = tk.Label(root, text="Select Protocol:")
protocol_label.pack(pady=10)

protocol_var = tk.StringVar(value="sftp")
protocol_dropdown = tk.OptionMenu(root, protocol_var, "sftp", "ftp")
protocol_dropdown.pack(pady=10)

# Create and place a button in the window
create_button = tk.Button(root, text="Create WinSCP Session", command=create_winscp_session)
create_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
