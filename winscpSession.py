import os
import tkinter as tk
from tkinter import messagebox

def save_winscp_session():
    # Define session details
    session_name = "MySession"
    host = 'example.com'
    port = '22'
    username = 'your_username'
    protocol = 'sftp'
    privatekey = 'C:\\path\\to\\private_key.ppk'
    hostkey = 'ssh-rsa 2048 ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    remotedir = '/home/your_username'
    localdir = 'C:\\Users\\your_user\\Documents'

    # Create a WinSCP script to save the session without connecting
    script_content = f"""
    open {protocol}://{username}@{host}:{port}/ -privatekey={privatekey} -hostkey="{hostkey}" -session={session_name}
    option batch abort
    option confirm off
    close
    save
    """

    # Save the script to a file
    script_filename = 'winscp_script.txt'
    with open(script_filename, 'w') as script_file:
        script_file.write(script_content)

    # Execute the WinSCP script via command line to save the session
    winscp_path = r'"C:\Program Files (x86)\WinSCP\winscp.com"'  # Adjust path if necessary
    os.system(f'{winscp_path} /script={script_filename}')

    # Show a message box to indicate the session has been saved
    messagebox.showinfo("WinSCP Session", "Session has been successfully saved without connecting!")

# Create the main application window
root = tk.Tk()
root.title("WinSCP Session Saver")

# Create and place a button in the window
save_button = tk.Button(root, text="Save WinSCP Session", command=save_winscp_session)
save_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
