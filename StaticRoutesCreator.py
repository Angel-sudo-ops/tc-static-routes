import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog
import xml.etree.ElementTree as ET
from xml.dom import minidom

default_file_path = os.path.join(r'C:\TwinCAT\3.1\Target', 'StaticRoutes.xml')

enableCC = False
# Function to create routes.xml with dynamic parameters
def create_routes_xml(project, lgv_list, base_ip, file_path, is_tc3):
    # Extract the base part of the IP address and the starting offset
    base_ip_parts = base_ip.rsplit('.', 1)
    base_ip_prefix = base_ip_parts[0]
    ip_offset = int(base_ip_parts[1])-lgv_list[0]

    # Create the root element
    config = ET.Element("TcConfig")
    config.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    routes = ET.SubElement(config,"RemoteConnections")

    # Create route elements dynamically
    for i in range(len(lgv_list)):
        current_offset = ip_offset + lgv_list[i]
        route_element = ET.SubElement(routes, "Route")
        
        name = ET.SubElement(route_element, "Name")
        if lgv_list[i] > 0 and lgv_list[i] < 10:
            name.text = f"CC{project}_LGV0{lgv_list[i]}"
        else:
            name.text = f"CC{project}_LGV{lgv_list[i]}"
        
        address = ET.SubElement(route_element, "Address")
        address.text = f"{base_ip_prefix}.{current_offset}"  # Increment IP offset
        
        netid = ET.SubElement(route_element, "NetId")
        netid.text = f"{base_ip_prefix}.{current_offset}.1.1"  # Increment NetId offset
        if is_tc3:
            netid.set("RemoteNetId", "192.168.11.2.1.1")

        type_ = ET.SubElement(route_element, "Type")
        type_.text = "TCP_IP"
        
        if is_tc3:
            timeout = ET.SubElement(route_element, "Flags")
            timeout.text = "32"

    # Create a tree from the root element and write it to a file in the pretty version
    # tree = ET.ElementTree(config)
    # tree.write(file_path, encoding="utf-8", xml_declaration=True)

    xmlstr = minidom.parseString(ET.tostring(config)).toprettyxml(indent="   ")
    #pretty_file_path = os.path.splitext(file_path)[0] + "_pretty.xml"
    with open(file_path, "w") as f:
        f.write(xmlstr)

    messagebox.showinfo("Success", "XML file has been created successfully!")

    global enableCC
    enableCC = True

def convert_static_to_cc(static_routes_file, CC_file):
    # Parse the StaticRoutes.xml file
    tree = ET.parse(static_routes_file)
    root = tree.getroot()
    
    # Create the root for cc.xml
    CC_root = ET.Element("Fleet")
    
    # Iterate through each route in the static routes file
    for route in root.find("RemoteConnections").findall("Route"):
        lgv = ET.SubElement(CC_root, "LGV")
        
        # Convert Name to Number (assumes last two digits are the number)
        name = route.find("Name").text
        number = name[-2:]
        ET.SubElement(lgv, "Number").text = str(int(number))  # Convert to int to remove leading zeroes
        
        # Add Type (fixed as "ELE" based on the example)
        ET.SubElement(lgv, "Type").text = "undefined"
        
        # Address to IP
        address = route.find("Address").text
        ET.SubElement(lgv, "IP").text = address
        
        # Netid to AMS
        netid = route.find("NetId").text
        ET.SubElement(lgv, "AMS").text = netid
    
    # Pretty print the XML
    raw_string = ET.tostring(CC_root, 'utf-8')
    reparsed = minidom.parseString(raw_string)
    pretty_xml = reparsed.toprettyxml(indent="  ")
    
    # Write to the file
    with open(CC_file, 'w', encoding='utf-8') as f:
        f.write(pretty_xml)
    
    messagebox.showinfo("Success", "ControlCenter file has been created successfully!")

def validate_and_create_cc():
    static_file_path = entry_file_path.get().strip()
    path_to_save_file = filedialog.asksaveasfilename(
        initialfile="ControlCenter",
        defaultextension=".xml", 
        filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
    )
    if path_to_save_file is None:
        messagebox.showerror("Input Error", "Please provide path to save the file.")
        return
    convert_static_to_cc(static_file_path, path_to_save_file)

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

# def validate_limit(*args):
#     limit = entry_limit.get().strip()
#     if limit.isdigit() and int(limit) > 0:
#         entry_limit.config(bg='white')
#     else:
#         entry_limit.config(bg='yellow')

# def validate_offset_lgv(*args):
#     offset_lgv = entry_offsetLGV.get().strip()
#     if offset_lgv.isdigit() and int(offset_lgv) > 0:
#         entry_offsetLGV.config(bg='white')
#     else:
#         entry_offsetLGV.config(bg='yellow')

def validate_base_ip(*args):
    base_ip = entry_base_ip.get().strip()
    if validate_ip(base_ip):
        entry_base_ip.config(bg='white')
    else:
        entry_base_ip.config(bg='yellow')

# Function to select the file save location
def select_file_path():
    file_path = filedialog.asksaveasfilename(
        initialfile="StaticRoutes",
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

def toggle_cc():
    if optionTC.get() == "TC3":
        if enableCC:
            create_cc.config(state='normal')
        else:
            create_cc.config(state='disabled')
    else:
        create_cc.config(state='disabled')

def create_placeholder(entry, placeholder_text):
    entry.insert(0, placeholder_text)
    entry.bind("<FocusIn>", lambda event: on_focus_in(entry, placeholder_text))
    entry.bind("<FocusOut>", lambda event: on_focus_out(entry, placeholder_text))

def on_focus_in(entry, placeholder_text):
    if entry.get() == placeholder_text:
        entry.delete(0, tk.END)
        entry.config(fg='black')

def on_focus_out(entry, placeholder_text):
    if not entry.get():
        entry.insert(0, placeholder_text)
        entry.config(fg='grey')

# Function to validate the inputs and create XML
def validate_and_create_xml():
    try:
        project = str(entry_project.get())
        if len(project) != 4:
            raise ValueError("Project number must be a 4 digit number")
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
    
    is_tc3 = optionTC.get() == "TC3"

    lgv_list = []
    range_input = entry_lgv_range.get()

    if not range_input:
        return None
    else:
        ranges = range_input.split(',')
        for r in ranges:
            if '-' in r:
                start, end = map(int, r.split('-'))
                lgv_list.extend(range(start, end + 1))
            else:
                lgv_list.append(r)
    
    create_routes_xml(project, lgv_list, base_ip, file_path, is_tc3)
    toggle_cc()

# Set up the GUI
root = tk.Tk()
root.title("Static Routes XML Creator")

#Disable resizing
root.resizable(False, False)

optionTC = tk.StringVar(value="TC3")
radio1 = tk.Radiobutton(root, text="TC2", variable=optionTC, value="TC2", command=toggle_cc)
radio1.grid(row=0, column=1, padx=5, pady=5, sticky='w')
radio2 = tk.Radiobutton(root, text="TC3", variable=optionTC, value="TC3", command=toggle_cc)
radio2.grid(row=0, column=1, padx=50, pady=5, sticky='w')

tk.Label(root, text="Project number CC:").grid(row=1, column=0, padx=10, pady=5)
entry_project = tk.Entry(root, fg="grey")
create_placeholder(entry_project, "e.g., 1584")
entry_project.grid(row=1, column=1, padx=10, pady=5)
entry_project.bind("<KeyRelease>", validate_project)

tk.Label(root, text="LGV numbers:").grid(row=2, column=0, padx=10, pady=5)
entry_lgv_range = tk.Entry(root, fg="grey")
create_placeholder(entry_lgv_range, "e.g., 1-5,11-17,20-25")
entry_lgv_range.grid(row=2, column=1, padx=10, pady=5)
# entry_lgv_range.bind("<KeyRelease>", validate_limit)

# tk.Label(root, text="Starting LGV:").grid(row=3, column=0, padx=10, pady=5)
# entry_offsetLGV = tk.Entry(root)
# entry_offsetLGV.grid(row=3, column=1, padx=10, pady=5)
# entry_offsetLGV.bind("<KeyRelease>", validate_offset_lgv)

tk.Label(root, text="First IP: ").grid(row=4, column=0, padx=10, pady=5)
entry_base_ip = tk.Entry(root, fg="grey")
create_placeholder(entry_base_ip, "e.g., 172.20.3.10")
entry_base_ip.grid(row=4, column=1, padx=10, pady=5)
entry_base_ip.bind("<KeyRelease>", validate_base_ip)

# File Path entry will be disabled and set dynamically
tk.Label(root, text="File Path:").grid(row=5, column=0, padx=10, pady=5)
#file_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'StaticRoutes.xml')  # Adjust folder as needed (e.g., 'Documents')
entry_file_path = tk.Entry(root, state='normal')
entry_file_path.grid(row=5, column=1, padx=10, pady=5)
entry_file_path.insert(0, default_file_path)
entry_file_path.config(state='disabled')

# Checkbox to toggle file path selection
select_path_var = tk.BooleanVar()
check_select_path = tk.Checkbutton(root, text="Select File Path", variable=select_path_var, command=toggle_file_path_selection)
check_select_path.grid(row=6, column=1, columnspan=2, pady=5)

# Button to select file path
button_select_path = tk.Button(root, text="Browse...", command=select_file_path)
button_select_path.grid(row=7, column=1, padx=10, pady=5)
button_select_path.config(state='disabled')

# Button to create Control Center XML
create_cc = tk.Button(root, text="Create CC XML", command=validate_and_create_cc)
create_cc.grid(row=7, column=0, pady=10)
create_cc.config(state='disable')

# Button to create XML
create_xml = tk.Button(root, text="Create XML", command=validate_and_create_xml)
create_xml.grid(row=8, columnspan=2, pady=10)

root.mainloop()

# hacer que solo pueda crear las rutas de CC cespués de haber hecho las rutas del StaticRoutes.xml - Done

# hacer que hagas todas las selecciones primero, rutas de TC2, rutas de TC3, cargadores, luego ya al final hacer el archivo

# ver si una vez que se crea el archivo, ver si se pueden ir agregando rutas al archivo existente