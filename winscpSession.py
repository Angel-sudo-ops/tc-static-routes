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
    privatekey = 'C:\\path\\to\\private_key.ppk' if protocol == 'sftp' else ''
    hostkey = 'ssh-rsa 2048 ABCDEFGHIJKLMNOPQRSTUVWXYZ' if protocol == 'sftp' else ''
    remotedir = '/home/your_username'
    localdir = 'C:\\Users\\your_user\\Documents'

    # Construct the command to save the session
    winscp_path = r'"C:\Program Files (x86)\WinSCP\winscp.com"'  # Adjust path if necessary
    session_command = f'{winscp_path} /log=winscp.log /defaults'

    if protocol == 'sftp':
        session_command += f' open sftp://{username}@{host}:{port}/ -privatekey="{privatekey}" -hostkey="{hostkey}"'
    else:
        session_command += f' open ftp://{username}@{host}:{port}/'

    # Save the session without connecting
    session_command += f' -sessionname={session_name}'
    os.system(session_command)

    # Show a message box to indicate the session has been saved
    messagebox.showinfo("WinSCP Session", f"{protocol.upper()} session has been successfully saved without connecting!")

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
