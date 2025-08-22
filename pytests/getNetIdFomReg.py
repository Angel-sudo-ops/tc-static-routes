import winreg
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

def get_ams_net_id():
    paths = [
        r"SOFTWARE\Beckhoff\TwinCAT3\System",  # TwinCAT 3
        r"SOFTWARE\Beckhoff\TwinCAT\System",    # TwinCAT 2
        r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT3\System",  # TwinCAT 3
        r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT\System"    # TwinCAT 2
    ]
    
    for path in paths:
        try:
            # Attempt to open the registry key
            reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
            ams_net_id, _ = winreg.QueryValueEx(reg_key, "AmsNetId")
            return ams_net_id
        except FileNotFoundError:
            continue  # If this path doesn't exist, try the next one
        except Exception as e:
            print(f"Failed to get AMS Net ID from registry path {path}: {e}")
            return None
    
    print("AMS Net ID not found in the registry.")
    return None

def format_ams_net_id(raw_net_id):
    # Convert the byte array to an IP-like string format
    return '.'.join(str(byte) for byte in raw_net_id)

def on_button_click():
    # Retrieve the AMS Net ID
    raw_ams_net_id = get_ams_net_id()
    
    if raw_ams_net_id:
        formatted_ams_net_id = format_ams_net_id(raw_ams_net_id)
        ams_net_id_var.set(formatted_ams_net_id)
    else:
        messagebox.showerror("Error", "Failed to retrieve the AMS Net ID.")
        ams_net_id_var.set("")

# Create the Tkinter window
root = tk.Tk()
root.title("NetID Finder")
root.geometry("250x150")

# Label for AMS Net ID
ttk.Label(root, text="Local AMS Net ID:").pack(pady=10)

# StringVar to hold the AMS Net ID
ams_net_id_var = tk.StringVar()

# Entry field to display the AMS Net ID
ams_net_id_entry = ttk.Entry(root, textvariable=ams_net_id_var, state="readonly")
ams_net_id_entry.pack(pady=10)

# Button to trigger AMS Net ID retrieval
ttk.Button(root, text="Get AMS Net ID", command=on_button_click).pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
