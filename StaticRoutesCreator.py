import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog
import xml.etree.ElementTree as ET

default_file_path = os.path.join(r'C:\TwinCAT\3.1\Target', 'StaticRoutesTest.xml')

# Function to create routes.xml with dynamic parameters
def create_routes_xml(project, limit, offset_lgv, base_ip, file_path):
    # Extract the base part of the IP address and the starting offset
    base_ip_parts = base_ip.rsplit('.', 1)
    base_ip_prefix = base_ip_parts[0]
    ip_offset = int(base_ip_parts[1])

    # Create the root element
    config = ET.Element("TcConfig")
    config.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    routes = ET.SubElement(config,"RemoteConnections")

    # Create route elements dynamically
    for i in range(limit):
        current_offset = ip_offset + i
        route_element = ET.SubElement(routes, "Route")
        
        name = ET.SubElement(route_element, "Name")
        if offset_lgv + i > 0 and offset_lgv + i < 10:
            name.text = f"CC{project}_LGV0{offset_lgv + i}"
        else:
            name.text = f"CC{project}_LGV{offset_lgv + i}"
        
        address = ET.SubElement(route_element, "Address")
        address.text = f"{base_ip_prefix}.{current_offset}"  # Increment IP offset
        
        netid = ET.SubElement(route_element, "NetId")
        netid.text = f"{base_ip_prefix}.{current_offset}.1.1"  # Increment NetId offset
        netid.set("RemoteNetId", "192.168.11.2.1.1")

        type_ = ET.SubElement(route_element, "Type")
        type_.text = "TCP_IP"
        
        timeout = ET.SubElement(route_element, "Flags")
        timeout.text = "32"

    # Create a tree from the root element and write it to a file in the pretty version
    # tree = ET.ElementTree(config)
    # tree.write(file_path, encoding="utf-8", xml_declaration=True)

    # Pretty-print the XML
    from xml.dom import minidom

    xmlstr = minidom.parseString(ET.tostring(config)).toprettyxml(indent="   ")
    #pretty_file_path = os.path.splitext(file_path)[0] + "_pretty.xml"
    with open(file_path, "w") as f:
        f.write(xmlstr)

    messagebox.showinfo("Success", "XML file has been created successfully!")

# Function to validate IP address format
def validate_ip(ip):
    # Compile the regex pattern for the base IP format 'xxx.xxx.xxx.xxx'
    pattern = re.compile(r'^(\d{1,3}\.){3}\d{1,3}$')  # Match format 'xxx.xxx.xxx.xxx'
    # Check if the input matches the pattern
    if pattern.match(ip):
        # Split the IP into parts and check if each part is between 0 and 255
        parts = ip.split('.')
        return all(0 <= int(num) <= 255 for num in parts)
    return False

# Real-time validation functions
def validate_project(*args):
    project = entry_project.get().strip()
    if project.isdigit() and len(project) == 4:
        entry_project.config(bg='white')
    else:
        entry_project.config(bg='yellow')

def validate_limit(*args):
    limit = entry_limit.get().strip()
    if limit.isdigit() and int(limit) > 0:
        entry_limit.config(bg='white')
    else:
        entry_limit.config(bg='yellow')

def validate_offset_lgv(*args):
    offset_lgv = entry_offsetLGV.get().strip()
    if offset_lgv.isdigit() and int(offset_lgv) > 0:
        entry_offsetLGV.config(bg='white')
    else:
        entry_offsetLGV.config(bg='yellow')

def validate_base_ip(*args):
    base_ip = entry_base_ip.get().strip()
    if validate_ip(base_ip):
        entry_base_ip.config(bg='white')
    else:
        entry_base_ip.config(bg='yellow')

# Function to select the file save location
def select_file_path():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xml", 
        filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
    )
    if file_path:
        entry_file_path.config(state='normal')
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)
        entry_file_path.config(state='disabled')

# Function to toggle file path selection
def toggle_file_path_selection():
    if select_path_var.get():
        button_select_path.config(state='normal')
        entry_file_path.config(state='normal')
        entry_file_path.delete(0, tk.END)
        entry_file_path.config(state='disabled')
    else:
        button_select_path.config(state='disabled')
        entry_file_path.config(state='normal')
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, default_file_path)
        entry_file_path.config(state='disabled')

# Function to validate the inputs and create XML
def validate_and_create_xml():
    try:
        project = str(entry_project.get())
        if len(project) != 4:
            raise ValueError("Project number must be a 4 digit number")
    except ValueError as e:
        messagebox.showerror("Invalid input", str(e))
        return
    try:
        limit = int(entry_limit.get())
        if limit <= 0:
            raise ValueError("The number of routes must be a positive integer.")
    except ValueError as e:
        messagebox.showerror("Invalid input", str(e))
        return
    try:
        offset_lgv = int(entry_offsetLGV.get())
        if offset_lgv <= 0:
            raise ValueError("Must be a positive integer.")
    except ValueError as e:
        messagebox.showerror("Invalid input", str(e))
        return

    base_ip = entry_base_ip.get().strip()
    if not validate_ip(base_ip):
        messagebox.showerror("Invalid input", "Please enter a valid base IP address in the format 'xxx.xxx.xxx.xxx'")
        return
    
    file_path = entry_file_path.get().strip()
    if not file_path:
        messagebox.showerror("Invalid input", "Please select a file path to save the XML.")
        return
    # if not validate_ip(base_ip):
    #      messagebox.showerror("Invalid input", "Please enter a valid base IP address in the format 'xxx.xxx.xxx'")
    #      return
    

    # type_value = entry_type.get().strip()
    # if not type_value:
    #     messagebox.showerror("Invalid input", "Type value cannot be empty.")
    #     return

    # timeout_value = entry_timeout.get().strip()
    # if not timeout_value.isdigit():
    #     messagebox.showerror("Invalid input", "Please enter a valid timeout value.")
    #     return

    # Construct file path dynamically using current user's home directory
    # username = os.getlogin()  # Get the current user's login name
    # file_name = "StaticRoutes.xml"
    # file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name)  # Adjust folder as needed (e.g., 'Documents')

    create_routes_xml(project, limit, offset_lgv, base_ip, file_path)

# Set up the GUI
root = tk.Tk()
root.title("Static Routes XML Creator")

#Disable resizing
root.resizable(False, False)

tk.Label(root, text="Project number CC:").grid(row=0, column=0, padx=10, pady=5)
entry_project = tk.Entry(root)
entry_project.grid(row=0, column=1, padx=10, pady=5)
entry_project.bind("<KeyRelease>", validate_project)

tk.Label(root, text="Number of LGVs:").grid(row=1, column=0, padx=10, pady=5)
entry_limit = tk.Entry(root)
entry_limit.grid(row=1, column=1, padx=10, pady=5)
entry_limit.bind("<KeyRelease>", validate_limit)

tk.Label(root, text="Starting LGV:").grid(row=2, column=0, padx=10, pady=5)
entry_offsetLGV = tk.Entry(root)
entry_offsetLGV.grid(row=2, column=1, padx=10, pady=5)
entry_offsetLGV.bind("<KeyRelease>", validate_offset_lgv)

tk.Label(root, text="Base IP Address (e.g., 172.20.3.10):").grid(row=3, column=0, padx=10, pady=5)
entry_base_ip = tk.Entry(root)
entry_base_ip.grid(row=3, column=1, padx=10, pady=5)
entry_base_ip.bind("<KeyRelease>", validate_base_ip)

# Type entry will be disabled
# tk.Label(root, text="Type:").grid(row=4, column=0, padx=10, pady=5)
# entry_type = tk.Entry(root, state='normal')
# entry_type.grid(row=4, column=1, padx=10, pady=5)
# entry_type.insert(0, "TCP_IP")
# entry_type.config(state='disabled')

# Flags entry will be disabled
# tk.Label(root, text="Flags:").grid(row=5, column=0, padx=10, pady=5)
# entry_flags = tk.Entry(root, state='normal')
# entry_flags.grid(row=5, column=1, padx=10, pady=5)
# entry_flags.insert(0, "32")
# entry_flags.config(state='disabled')

# File Path entry will be disabled and set dynamically
tk.Label(root, text="File Path:").grid(row=6, column=0, padx=10, pady=5)
#file_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'StaticRoutes.xml')  # Adjust folder as needed (e.g., 'Documents')
entry_file_path = tk.Entry(root, state='normal')
entry_file_path.grid(row=6, column=1, padx=10, pady=5)
entry_file_path.insert(0, default_file_path)
entry_file_path.config(state='disabled')

# Checkbox to toggle file path selection
select_path_var = tk.BooleanVar()
check_select_path = tk.Checkbutton(root, text="Select File Path", variable=select_path_var, command=toggle_file_path_selection)
check_select_path.grid(row=7, columnspan=2, pady=5)

# Button to select file path
button_select_path = tk.Button(root, text="Browse...", command=select_file_path)
button_select_path.grid(row=8, column=1, padx=10, pady=5)
button_select_path.config(state='disabled')

# Button to create XML
tk.Button(root, text="Create XML", command=validate_and_create_xml).grid(row=9, columnspan=2, pady=10)

root.mainloop()
