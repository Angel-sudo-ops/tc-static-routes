import configparser
import os
import tkinter as tk
from tkinter import messagebox

def save_winscp_session():
    # Get the selected protocol
    protocol = protocol_var.get()

    # Define session details
    session_name = "MySession"
    host = 'example.com'
    port = '22' if protocol == 'sftp' else '21'
    username = 'your_username'
    localdir = 'C:\\Users\\your_user\\Documents'
    remotedir = '/home/your_username'

    # Path to the winscp.ini file
    ini_path = os.path.expanduser('~\\AppData\\Roaming\\WinSCP.ini')

    # Create config parser and read the INI file
    config = configparser.ConfigParser()
    config.read(ini_path)

    # Create a new session section in the INI file
    section_name = f'Session.{session_name}'
    config[section_name] = {
        'HostName': host,
        'UserName': username,
        'PortNumber': port,
        'FSProtocol': '2' if protocol == 'sftp' else '0',
        'LocalDirectory': localdir,
        'RemoteDirectory': remotedir
    }

    # Save the session to the INI file
    with open(ini_path, 'w') as configfile:
        config.write(configfile)

    # Show a message box to indicate the session has been saved
    messagebox.showinfo("WinSCP Session", f"{protocol.upper()} session has been successfully saved in {ini_path}!")

# Create the main application window
root = tk.Tk()
root.title("WinSCP Session Saver")

# Create a label and dropdown for protocol selection
protocol_label = tk.Label(root, text="Select Protocol:")
protocol_label.pack(pady=10)

protocol_var = tk.StringVar(value="sftp")
protocol_dropdown = tk.OptionMenu(root, protocol_var, "sftp", "ftp")
protocol_dropdown.pack(pady=10)

# Create and place a button in the window
save_button = tk.Button(root, text="Save WinSCP Session", command=save_winscp_session)
save_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
