import sys
import os
import platform
import threading
import re
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import xml.etree.ElementTree as ET
from xml.dom import minidom
import sqlite3
import winreg as reg
import configparser
import clr
import System
from System import Activator 
from System import Type
from System.Reflection import BindingFlags
from System.Net import IPAddress

__version__ = '1.7'

default_file_path = os.path.join(r'C:\TwinCAT\3.1\Target', 'StaticRoutes.xml')

class ToolTip:
    def __init__(self, widget, text, delay=500, fade_duration=500):
        self.widget = widget
        self.text = text
        self.delay = delay  # delay before showing tooltip in milliseconds
        self.fade_duration = fade_duration  # duration of fade effect in milliseconds
        self.tooltip_window = None
        self.id = None
        self.opacity = 0
        self.is_fading_out = False

        self.widget.bind("<Enter>", self.schedule_tooltip)
        self.widget.bind("<Leave>", self.start_fade_out)
        self.widget.bind("<Button-1>", self.on_click)
        self.widget.winfo_toplevel().bind("<Motion>", self.check_motion)

    def schedule_tooltip(self, event):
        self.cancel_tooltip()
        if not self.is_fading_out:
            self.id = self.widget.after(self.delay, self.show_tooltip)

    def show_tooltip(self):
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes('-alpha', 0.0)  # Start with full transparency

        label = tk.Label(tw, text=self.text, justify='left',
                         background="white", relief='solid', borderwidth=1,
                         font=("helvetica", "8", "normal"))
        label.pack(ipadx=1)

        self.is_fading_out = False
        self.fade_in()

    def fade_in(self):
        if self.opacity < 1.0 and not self.is_fading_out:
            self.opacity += 0.05
            self.tooltip_window.attributes('-alpha', self.opacity)
            self.tooltip_window.after(int(self.fade_duration / 20), self.fade_in)
        else:
            if self.opacity >= 1.0:
                self.opacity = 1.0

    def start_fade_out(self, event=None):
        if self.tooltip_window and not self.is_fading_out:
            self.is_fading_out = True
            self.fade_out()

    def fade_out(self):
        if self.opacity > 0:
            self.opacity -= 0.05
            if self.tooltip_window:
                self.tooltip_window.attributes('-alpha', self.opacity)
                self.tooltip_window.after(int(self.fade_duration / 20), self.fade_out)
        else:
            if self.tooltip_window:
                self.tooltip_window.destroy()
                self.tooltip_window = None
                self.opacity = 0  # Reset opacity for the next tooltip
            self.is_fading_out = False

    def cancel_tooltip(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def check_motion(self, event):
        widget_under_cursor = self.widget.winfo_containing(event.x_root, event.y_root)
        if widget_under_cursor != self.widget and self.tooltip_window and not self.is_fading_out:
            self.start_fade_out()
    
    def on_click(self, event):
        # Reset the tooltip logic on click to ensure it can still appear
        self.cancel_tooltip()
        if self.tooltip_window:
            self.start_fade_out()
        self.schedule_tooltip(event)

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
    # static_file_path = entry_file_path.get().strip()
    static_file_path = default_file_path

    path_to_save_file = filedialog.asksaveasfilename(
        initialdir= os.path.join(os.path.expanduser("~"), "Documents"),
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

def validate_ams_net_id(ams_net_id):
    # Regular expression that checks for the general IP format and ends with .1.1
    pattern = re.compile(r'^(\d{1,3}\.){3}\d{1,3}\.1\.1$')
    if pattern.match(ams_net_id):
        parts = ams_net_id.split('.')
        # Ensure each segment before .1.1 is an integer between 0 and 255
        return all(0 <= int(num) <= 255 for num in parts[:-2])
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

######################################## placeholders #######################################33
placeholders = {}

def create_placeholder(entry, placeholder_text):
    entry.insert(0, placeholder_text)
    entry.bind("<FocusIn>", lambda event: on_focus_in(entry, placeholder_text))
    entry.bind("<FocusOut>", lambda event: on_focus_out(entry, placeholder_text))
    placeholders[entry] = placeholder_text

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

################################### Get StaticRoutes.xml and create table #########################

def populate_table_from_xml():
    # Ask the user to select an XML file
    file_path = filedialog.askopenfilename(title="Select StaticRoutes file", 
                                           initialdir="C:\\TwinCAT\\3.1\\Target",
                                           filetypes=[("XML files", "*.xml")])
    
    if file_path:
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
        except ET.ParseError:
            messagebox.showerror("Error", "The selected file is not a valid XML file.")
            return

        # Check for the expected root elements
        remote_connections = root.find('RemoteConnections')
        if remote_connections is None:
            messagebox.showerror("Error", "XML file does not contain the expected 'RemoteConnections' structure.")
            return
        
        # Clear the existing table data
        for i in treeview.get_children():
            treeview.delete(i)
        
        # Initialize an empty list to hold the data
        routes_data = []
        
        # Iterate through each <Route> element in the XML
        for route in remote_connections.findall('Route'):
            name = route.find('Name')
            address = route.find('Address')
            net_id = route.find('NetId')

            if None in (name, address, net_id):
                messagebox.showwarning("Warning", "One or more routes are missing required fields (Name, Address, NetId).")
                continue  # Skip this route and move to the next

            name = name.text
            address = address.text
            net_id = net_id.text
            type_tc = "TC3" if route.find('Flags') is not None else "TC2"
            
            # Append the tuple to the list
            routes_data.append((name, address, net_id, type_tc))
        
        # Populate the Treeview with the data
        for item in routes_data:
            treeview.insert("", "end", values=item)
        # messagebox.showinfo("Success", "Data loaded successfully from the XML file.")

############################# Populate table based on inputs ##############################

def populate_table_from_inputs():
    project = entry_project.get()
    lgv_range = entry_lgv_range.get()
    base_ip = entry_base_ip.get()
    is_tc3 = optionTC.get() == "TC3"
    # is_lgv = optionLGV.get() == "LGV"

    try:
        if len(project) != 4 or project == placeholders[entry_project]:
            raise ValueError("Project number must be a 4 digit number")
        
    except ValueError as e:
        messagebox.showerror("Invalid input", str(e))
        return
    
    if lgv_range is None or lgv_range == placeholders[entry_lgv_range]:
        messagebox.showerror("Invalid input", "Please enter a valid range")
        return
    
    if not validate_ip(base_ip) or base_ip == placeholders[entry_base_ip]:
        messagebox.showerror("Invalid input", "Please enter a valid base IP address in the format 'xxx.xxx.xxx.xxx'")
        return
    
    lgvs = parse_range(lgv_range)

    # Parse the IPs based on the given base IP and LGV list
    ip_list = parse_ip(base_ip, lgvs)

    # Clear existing table data
    # for i in treeview.get_children():
    #     treeview.delete(i)

    # Loop through the parsed IPs and add to the table
    for i, current_ip in enumerate(ip_list):
        net_id = f"{current_ip}.1.1"
        route_name = f"CC{project}_{optionLGV.get()}{str(lgvs[i]).zfill(2)}"
        
        # Check if the record already exists in the table, ignoring the TC type and name
        record_exists = False
        for row in treeview.get_children():
            existing_values = treeview.item(row)["values"]
            if (existing_values[1] == current_ip and 
                existing_values[2] == net_id):
                record_exists = True
                # messagebox.showwarning("Duplicate Entry", f"The IP {current_ip} already exists.")

                break

        # Only add the record if it doesn't already exist
        if not record_exists:
            treeview.insert("", "end", values=(
                route_name,
                current_ip,
                net_id,
                "TC3" if is_tc3 else "TC2"
            ))

##################### Function to check for duplicates in the Treeview ######################
#Check for duplicates when input is the entry fields
# def is_duplicate(name, address, netid, type_tc):
#     for item in treeview.get_children():
#         existing_values = treeview.item(item, 'values')
#         if (name, address, netid, type_tc) == existing_values:
#             return True
#     return False

#Check for duplicates when input is the table directly
def is_duplicate(col_index, new_value, current_row):
    # Check for duplicates in the column except for the current editing row
    for item in treeview.get_children():
        if item != current_row:
            if treeview.item(item, 'values')[col_index] == new_value:
                return True
    return False

############################### Get IPs ###################################################
def parse_ip(base_ip, lgv_list):
    base_ip_prefix, start_ip = base_ip.rsplit('.', 1)
    start_ip = int(start_ip) - lgv_list[0]
    
    ip_list = []
    for lgv in lgv_list:
        current_ip = f"{base_ip_prefix}.{start_ip + lgv}"
        ip_list.append(current_ip)
    
    return ip_list

############################ Get range ##################################################
def parse_range(range_str):
    print(range_str)
    lgv_list = []
    try:
        for part in range_str.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                lgv_list.extend(range(start, end + 1))
            else:
                lgv_list.append(int(part))
        return lgv_list
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter a valid integer")

############################### Delete selected record ####################################
# With a button
def delete_selected():
    selected_item = treeview.selection()
    if selected_item:
        treeview.delete(selected_item)
    else:
        messagebox.showwarning("Selection Error", "Please select a record to delete.")

# With right click
def delete_selected_record_from_menu():
    selected_item = treeview.selection()
    if selected_item:
        treeview.delete(selected_item)

# Function to show the context menu
def show_context_menu(event):
    # Check if a record is selected
    selected_item = treeview.identify_row(event.y)
    if selected_item:
        treeview.selection_set(selected_item)
        context_menu.post(event.x_root, event.y_root)

# With DEL key
def delete_selected_record(event):
    selected_item = treeview.selection()
    if selected_item:
        treeview.delete(selected_item)

################################### Delete whole table ########################################

def delete_whole_table():
    for i in treeview.get_children():
        treeview.delete(i)

################################## Create StaticRoutes.xml from table ##########################

def get_table_data():
    rows = []
    for item in treeview.get_children():
        rows.append(treeview.item(item)["values"])
    print(rows)
    return rows
    
def create_routes_xml_from_table(file_path):
    data = get_table_data()
    # Create the root element
    config = ET.Element("TcConfig")
    config.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    # Create the RemoteConnections element
    routes = ET.SubElement(config, "RemoteConnections")

    # Iterate over the data to create the Route elements
    for row in data:
        name, address, netid, tc_type = row

        route_element = ET.SubElement(routes, "Route")

        ET.SubElement(route_element, "Name").text = name
        ET.SubElement(route_element, "Address").text = address
        
        netid_element = ET.SubElement(route_element, "NetId")
        netid_element.text = netid
        
        ET.SubElement(route_element, "Type").text = "TCP_IP"

        if tc_type == "TC3":
            netid_element.set("RemoteNetId", "192.168.11.2.1.1")
            ET.SubElement(route_element, "Flags").text = "32"

    # Convert to a pretty XML string
    xmlstr = minidom.parseString(ET.tostring(config, 'utf-8')).toprettyxml(indent="   ")

    # Write to a file
    with open(file_path, "w", encoding='utf-8') as f:
        f.write(xmlstr)

    messagebox.showinfo("Success", "StaticRoutes file has been created successfully!")

def save_routes_xml():
    if get_table_data() == []:
        messagebox.showerror("Attention", "Routes table is empty!")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xml",
                                             initialdir="C:\\TwinCAT\\3.1\\Target",
                                             initialfile="StaticRoutes.xml",
                                             filetypes=[("XML files", "*.xml")])
    if file_path:
        create_routes_xml_from_table(file_path)

######################################## Create Control Center xml file from table ################################
def create_cc_xml_from_table(file_path):
    data = get_table_data()

    # Create the root element
    fleet = ET.Element("Fleet")

    # Iterate over the data to create the Route elements
    for row in data:
        name, address, netid, tc_type = row
        if tc_type == 'TC3':
            lgv = ET.SubElement(fleet, "LGV")

            number = name[-2:]
            ET.SubElement(lgv, "Number").text = str(int(number))  # Convert to int to remove leading zeroes

            ET.SubElement(lgv, "Type").text = "undef"

            ip_element = ET.SubElement(lgv, "IP")
            ip_element.text = address

            netid_element = ET.SubElement(lgv, "AMS")
            netid_element.text = netid

    # Convert to a pretty XML string
    xmlstr = minidom.parseString(ET.tostring(fleet, 'utf-8')).toprettyxml(indent="   ")

    # Write to a file
    with open(file_path, "w", encoding='utf-8') as f:
        f.write(xmlstr)

    messagebox.showinfo("Success", "ControlCenter file has been created successfully!")

def save_cc_xml():
    if get_table_data() == []:
        messagebox.showerror("Attention", "Routes table is empty!")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xml",
                                             initialdir= os.path.join(os.path.expanduser("~"), "Documents"),
                                             initialfile="ControlCenter.xml",
                                             filetypes=[("XML files", "*.xml")])
    if file_path:
        create_cc_xml_from_table(file_path)

######################################## Modify data directly on table ############################################

def on_double_click(event):
    region = treeview.identify("region", event.x, event.y)
    if region == "cell":
        column = treeview.identify_column(event.x)
        row = treeview.identify_row(event.y)
        col_index = int(column.replace("#", "")) - 1
        current_value = treeview.item(row, 'values')[col_index]

        if col_index == 3:  # Assuming column 3 is the Type column
            create_combobox_for_type(column, row)
        else:
            create_entry_for_editing(column, row, col_index, current_value)


def create_combobox_for_type(column, row):
    bbox = treeview.bbox(row, column)
    if not bbox:
        return
    
    combo_edit = ttk.Combobox(treeview, values=["TC2", "TC3"], state="readonly")
    x, y, width, height = treeview.bbox(row, column)
    combo_edit.place(x=x, y=y, width=width, height=height)

    def on_select(event):
        if combo_edit.winfo_exists():
            treeview.set(row, column=column, value=combo_edit.get())
            combo_edit.destroy()
    
    def check_focus(event):
        # Destroy the Combobox if it is not the focus
        if event.widget != combo_edit:
            combo_edit.destroy()

    combo_edit.bind("<<ComboboxSelected>>", on_select)
    root.bind("<Button-1>", check_focus, add="+")  # Use "+" to add to existing bindings
    # combo_edit.focus()


def create_entry_for_editing(column, row, col_index, current_value):
    bbox = treeview.bbox(row, column)
    if not bbox:
        return
    
    entry_edit = tk.Entry(treeview, border=0)
    entry_edit.insert(0, current_value)
    x, y, width, height = treeview.bbox(row, column)
    entry_edit.place(x=x, y=y, width=width, height=height)
    entry_edit.focus()
    entry_edit.select_range(0, tk.END)

    def save_edit(event):
        if entry_edit.winfo_exists():
            new_value = entry_edit.get()
            if is_duplicate(col_index, new_value, row):
                messagebox.showerror("Invalid Input", f"Duplicate value found for {treeview.heading(col_index, 'text')}.")
                return  # Do not destroy the Entry, allow user to correct it
            if col_index == 0:  # Assuming the "Name" column is the first column (index 0)
                if not new_value.strip():  # Check if the name is not empty
                    messagebox.showerror("Invalid Input", "Name field cannot be empty.")
                    return # Do not destroy the Entry, allow user to correct it
                entry_edit.destroy() # Only destroy if validation is passed or not needed
            if (col_index == 1 and not validate_ip(new_value)) or \
                (col_index == 2 and not validate_ams_net_id(new_value)):
                messagebox.showerror("Invalid Input", "Please enter a valid value.")
                return
            entry_edit.destroy()
            
            treeview.set(row, column=column, value=new_value)

    def cancel_edit(event=None):
        if entry_edit.winfo_exists():
            entry_edit.destroy()

    entry_edit.bind("<Return>", save_edit)
    entry_edit.bind("<Escape>", lambda e: cancel_edit())
    entry_edit.bind("<FocusOut>", lambda e: cancel_edit())


################################## Sorting ################################################
def setup_treeview():
    # Initialize the headings with custom names
    headings = {
        'Name': 'Route Name',
        'Address': 'IP Address',
        'NetId': 'AMS Net Id',
        'Type': 'Type'
    }
    
    for col in treeview['columns']:
        treeview.heading(col, text=headings[col], command=lambda _col=col: treeview_sort_column(treeview, _col, False))

def treeview_sort_column(tv, col, reverse):
    # Retrieve all data from the treeview
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    # Sort the data
    l.sort(reverse=reverse, key=lambda t: natural_keys(t[0]))

    # Rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    # Dictionary to maintain custom headings
    headings = {
        'Name': 'Route Name',
        'Address': 'IP Address',
        'NetId': 'AMS Net Id',
        'Type': 'Type'
    }

    # Change the heading to show the sort direction
    for column in tv['columns']:
        heading_text = headings[column] + (' ↓' if reverse and column == col else ' ↑' if not reverse and column == col else '')
        tv.heading(column, text=heading_text, command=lambda _col=column: treeview_sort_column(tv, _col, not reverse))

def natural_keys(text):
    """
    Alphanumeric (natural) sort to handle numbers within strings correctly
    """
    return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]


#################################### Read config.db3 #######################################
def read_db3_file(db3_file_path, table_name):
    try:
        # Connect to the .db3 file
        conn = sqlite3.connect(db3_file_path)
        cursor = conn.cursor()

        # Check if the table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
        if not cursor.fetchone():
            # messagebox.showerror("Error", f"Table '{table_name}' does not exist in the database.")
            messagebox.showerror("Error", f"Wrong database format.")
            conn.close()
            return None, None

        # Query to get all rows from the specified table
        cursor.execute(f"SELECT * FROM {table_name}")
        
        # Fetch all rows
        rows = cursor.fetchall()

        # Get column names
        column_names = [description[0] for description in cursor.description]

        # Convert the rows into a list of dictionaries
        dict_rows = [dict(zip(column_names, row)) for row in rows]

        # Close the connection
        conn.close()

        return dict_rows
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return None

def populate_table_from_db3():
    project = entry_project.get()
    if len(project) != 4 or project == placeholders[entry_project]:
        messagebox.showinfo("Attention", "Add project number")
        return
    
    db3_path = filedialog.askopenfilename(title="Select config.db3 file", 
                                          initialdir="C:\\Program Files (x86)\\Elettric80",
                                          filetypes=[("DB3 files", "*.db3")])
    table_agvs = "tbl_AGVs"
    rows_agvs = read_db3_file(db3_path, table_agvs)

    table_param = "tbl_Parameter"
    rows_param = read_db3_file(db3_path, table_param)
    
    # # print(columns, rows)
    # for row in rows_agvs:
    #     if row['dbf_Enabled']:
    #         print(f"LGV{str(row['dbf_ID']).zfill(2)}")

    # Clear the existing table data
    for i in treeview.get_children():
        treeview.delete(i)
    
    # Default type_tc based on the transfer mode
    default_type_tc = "TC2"  # Assume TC2 unless specified otherwise
    for row_param in rows_param:
        if row_param['dbf_Name'] == "agvlayoutloadmethod" and row_param['dbf_Value'] == "SFTP":
            default_type_tc = "TC3" # If SFTP, set all to TC3

    # Initialize an empty list to hold the data
    routes_data = []
    # Iterate through each <Route> element in the XML
    for route in rows_agvs:
        if route['dbf_Enabled']: 
        # if None in (name, address, net_id):
        #     messagebox.showwarning("Warning", "One or more routes are missing required fields (Name, Address, NetId).")
        #     continue  # Skip this route and move to the next

            name = f"CC{project}_LGV{str(route['dbf_ID']).zfill(2)}"
            address = route['dbf_IP']
            net_id = f"{address}.1.1"
            
            # if route['Dbf_Comm_Library']>20 or 
            if route['LayoutCopy_Protocol']=="SFTP":
                type_tc = "TC3" 
            elif route['LayoutCopy_Protocol']=="FTP" or route['LayoutCopy_Protocol']=="NETFOLDER":
                type_tc = "TC2" 
            else:
                type_tc = default_type_tc 
        
            # Append the tuple to the list
            routes_data.append((name, address, net_id, type_tc))
    
    # Populate the Treeview with the data
    for item in routes_data:
        treeview.insert("", "end", values=item)

################################# Slipt project and LGV numer ######################################
def split_string(input_string):
    # Regular expression pattern for CCxxxxLGVxx or CCxxxx_LGVxx
    pattern_with_underscore = r"^CC\d{4}_(LGV|CB|BC)\d{2}$"
    pattern_without_underscore = r"^CC\d{4}(LGV|CB|BC)\d{2}$"
    
    # Check if the input string matches the expected patterns
    if re.match(pattern_with_underscore, input_string):
        # Split by underscore if present
        parts = input_string.split('_')
    elif re.match(pattern_without_underscore, input_string):
         # Use regular expression to split between the numeric and alphanumeric parts
        match = re.match(r"(CC\d{4})(LGV\d{2}|CB\d{2}|BC\d{2})", input_string)
        parts = [match.group(1), match.group(2)]
    else:
        # Raise a ValueError if the string does not match the expected format
        messagebox.showerror("Error", "At least one route name does not match the expected format: CCxxxxLGV/CB/BCxx or CCxxxx_LGV/CB/BCxx")
    
    return parts
################################### Create ini file for WinSCP connections ##########################
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
    
# Function to check if a HostName already exists in the INI file
def hostname_exists(config, host_name):
    for section in config.sections():
        if config.has_option(section, 'HostName') and config.get(section, 'HostName') == host_name:
            return True
    return False

# Function to create a session in the winscp.ini file
def create_winscp_ini_from_table(ini_path):
    # Ensure the directory for the INI file exists
    ini_dir = os.path.dirname(ini_path)
    os.makedirs(ini_dir, exist_ok=True)

    # Set the custom INI path in the registry
    if not set_custom_ini_path(ini_path):
        return "Failed to set the custom INI path in the registry."

    # Create config parser and read the INI file (if it exists)
    config = configparser.ConfigParser()
    if os.path.exists(ini_path):
        if config.has_section(r"Sessions\CC1548/LGV41"):
            config.remove_section('SshHostKeys')
        config.read(ini_path)
           
    
    data = get_table_data()
    repeated = 0
    total = 0
    for row in data:
        total = total + 1 
        name, address, netid, tc_type = row
        name_parts = split_string(str(name))
        folder_name = name_parts[0]
        session_name = name_parts[1]

        section_name = f'Sessions\\{folder_name}/{session_name}'

        # Check if the HostName already exists
        if hostname_exists(config, address):
            repeated=repeated+1
            print("HostName already exists.")
            continue 

        # Define session details
        config[section_name] = {
            'HostName': address,
            'PortNumber': '20022' if tc_type == "TC3" else '21',
            'UserName': 'Administrator' if tc_type == "TC3" else 'anonymous',
            'Password': 'A35C45504648113EE96A1003AC13A5A41D38313532352F282E3D28332E6D6B6E726E6C726E726A6A6D84CA5BFA50425E8C85' if tc_type == "TC3" else 'A35C755E6D593D323332253133292F6D6B6E726E6C726E72696D3D323332253133292F1C39243D312C3039723F333130FAB0',
        }

        # Add FSProtocol only for FTP
        if tc_type == "TC2":
            config[section_name]['FSProtocol'] = '5'


    # Write the session to the INI file
    with open(ini_path, 'w') as configfile:
        config.write(configfile)

    messagebox.showinfo("Success", f"Session created successfully in {ini_path} with {repeated} repeated routes out of {total}")

def save_winscp_ini():
    # Custom path for the INI file (not in Roaming)
    file_path = r'C:\WinSCPConfig\WinSCP.ini'  # Adjust this path if needed

    if get_table_data() == []:
        messagebox.showerror("Attention", "Routes table is empty!")
        return
    # file_path = filedialog.asksaveasfilename(defaultextension=".ini",
    #                                          initialdir= os.path.join(os.path.expanduser("~"), "Documents"),
    #                                          initialfile="WinSCP.ini",
    #                                          filetypes=[("INI files", "*.ini")])
    if file_path:
        create_winscp_ini_from_table(file_path)

######################################################## Create Routes ############################################################################

def load_dll():
    # Determine if the application is running as a standalone executable
    if getattr(sys, 'frozen', False):
        # If the application is frozen (bundled by PyInstaller), get the path of the executable
        base_path = sys._MEIPASS
    else:
        # If running as a script, use the current directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    # Construct the full path to the DLL
    dll_path = os.path.join(base_path, "CRADSDriver.dll")

    # Load the assembly
    clr.AddReference(dll_path)

def initialize_twincat_com():
    # Use the fully qualified name, including the assembly name, if necessary
    assembly_name = "CRADSDriver"
    type_name = "TwinCATAds.TwinCATCom, " + assembly_name

    # Get the type using the fully qualified name
    twincat_type = Type.GetType(type_name)

    if twincat_type is None:
        print(f"Failed to find type '{type_name}'")
        return None

    # Instantiate the TwinCATCom object using Activator
    twincat_com = Activator.CreateInstance(twincat_type)

    # Manually invoke CreateDLLInstance method to ensure initialization
    create_dll_instance_method = twincat_type.GetMethod(
        "CreateDLLInstance",
        BindingFlags.NonPublic | BindingFlags.Instance
    )
    if create_dll_instance_method:
        create_dll_instance_method.Invoke(twincat_com, None)
    else:
        print("Failed to locate CreateDLLInstance method")
        return None

    # Check the DLL dictionary and manually add an entry if necessary
    my_dll_instance_field = twincat_type.GetField("MyDLLInstance", BindingFlags.NonPublic | BindingFlags.Instance)
    dll_field = twincat_type.GetField("DLL", BindingFlags.NonPublic | BindingFlags.Static)
    
    if my_dll_instance_field and dll_field:
        my_dll_instance_value = my_dll_instance_field.GetValue(twincat_com)
        dll_dict = dll_field.GetValue(None)
        print(f"MyDLLInstance value after CreateDLLInstance: {my_dll_instance_value}")
        print(f"DLL Dictionary contains {len(dll_dict)} items after CreateDLLInstance")

        if my_dll_instance_value not in dll_dict:
            ads_for_twincat_ex_type = Type.GetType("TwinCATAds.ADSforTwinCATEx, " + assembly_name)
            if ads_for_twincat_ex_type:
                ads_for_twincat_ex = Activator.CreateInstance(ads_for_twincat_ex_type)
                # ADSforTwinCATEx is the type of values for the DLL dictionary
                dll_dict[my_dll_instance_value] = ads_for_twincat_ex 
                print(f"Manually added DLL entry for key {my_dll_instance_value}")
            else:
                print("Failed to create ADSforTwinCATEx instance")
                return None
            
    # Check the DLL dictionary after manual insertion
    if dll_field:
        dll_dict = dll_field.GetValue(None)
        print(f"DLL Dictionary contains {len(dll_dict)} items after manual insertion")

    return twincat_com

def get_local_ams_netid():
    # Use the fully qualified name, including the assembly name, if necessary
    assembly_name = "CRADSDriver"
    type_name = "TwinCATAds.ADSforTwinCAT, " + assembly_name

    # Get the type using the fully qualified name
    ads_twincat_type = Type.GetType(type_name)

    if ads_twincat_type is None:
        print(f"Failed to find type '{type_name}'")
        return None

    # Instantiate the TwinCATCom object using Activator
    ads_twincat = Activator.CreateInstance(ads_twincat_type)

    local_ams_netid = ads_twincat.get_MyAMSNetID()
    print(f"Local AMS NetID: {local_ams_netid}")

    return local_ams_netid

def create_route(twincat_com, entry, username, password, netid_ip, system_name):
    name, ip, amsnet_id, type_ = entry

    # Determine the port based on TC2 or TC3
    port = 851 if type_ == 'TC3' else 801

    # Set properties for the current IP
    twincat_com.DisableSubScriptions = True
    twincat_com.Password = password
    twincat_com.PollRateOverride = 500
    twincat_com.TargetAMSNetID = amsnet_id
    twincat_com.TargetIPAddress = ip
    twincat_com.TargetAMSPort = port
    twincat_com.UserName = username
    twincat_com.UseStaticRoute = True

    local_ip = IPAddress.Parse(netid_ip)

    # Call CreateRoute
    try:
        result = twincat_com.CreateRoute(system_name, local_ip)
        print(f"Route created successfully for {name} ({ip}), result: {result}\n")
    except Exception as e:
        print(f"Error during CreateRoute invocation for {name} ({ip}): {e}\n")

def create_routes_from_data(data, username, password):
    load_dll()
    # Initialize TwinCATCom only once
    twincat_com = initialize_twincat_com()
    if not twincat_com:
        print("Failed to initialize TwinCATCom")
        return
    # netid_ip = '10.230.0.34' #Replace this with a method to get AMS Net ID
    netid = str(get_local_ams_netid()).split('.')
    netid_ip = '.'.join(netid[:4])
    system_name = platform.node()
    print(f"IP: {netid_ip}")
    print(f"System Name: {system_name}")
    for entry in data:
        threading.Thread(target=create_route, args=(twincat_com, entry, username, password, netid_ip, system_name)).start()

def create_tc_routes():
    data = get_table_data()
    username = username_entry.get()
    password = password_entry.get()
    if username == "":
        messagebox.showerror("Attention", "Add username!")
        return
    if password == "":
        messagebox.showerror("Attention", "Add password!")
        return
    if data == []:
        messagebox.showerror("Attention", "Routes table is empty!")
        return
    create_routes_from_data(data, username, password)


################################### Button design ##########################################
def on_enter(e):
    if e.widget['state']== "normal":
        e.widget['background'] = 'LightSkyBlue1'

def on_leave(e):
    if e.widget['state'] == "normal":
        e.widget['background'] = 'ghost white'

def button_design(entry):
    entry.bind("<Enter>", on_enter)
    entry.bind("<Leave>", on_leave)
    entry.bind("<Button-1>", on_enter)


####################################### Set up the GUI ######################################
root = tk.Tk()
root.title(f"Static Routes XML Creator {__version__}")

# Check if running as a script or frozen executable
if getattr(sys, 'frozen', False):
    icon_path = os.path.join(sys._MEIPASS, "./route.ico")
else:
    icon_path = os.path.abspath("./route.ico")
root.iconbitmap(icon_path)

window_width = 470
window_lenght = 570
root.geometry(f"{window_width}x{window_lenght}")
root.minsize(window_width, window_lenght)
# root.resizable(True, True)

frame_tc = tk.Frame(root)
frame_tc.grid(row=0, column=1, padx=5, pady=5)
optionTC = tk.StringVar(value="TC3")
tc2_radio = tk.Radiobutton(frame_tc, text="TC2", variable=optionTC, value="TC2")
tc2_radio.grid(row=0, column=0, padx=0, pady=0, sticky='w')
tc3_radio = tk.Radiobutton(frame_tc, text="TC3", variable=optionTC, value="TC3")
tc3_radio.grid(row=0, column=1, padx=0, pady=0, sticky='w')

frame_project = tk.Frame(root)
frame_project.grid(row=0, column=0, padx=5, pady=5, sticky='e')
label_project = tk.Label(frame_project, text="Project number CC:")
label_project.grid(row=0, column=0, padx=5, pady=5)

entry_project = tk.Entry(frame_project, fg="grey")
entry_project.grid(row=0, column=1, padx=5, pady=5)
create_placeholder(entry_project, "e.g., 1584")
entry_project.bind("<KeyRelease>", validate_project)

frame_range = tk.Frame(root)
frame_range.grid(row=1, column=0, padx=5, pady=5, sticky='e')
frame_lgv = tk.Frame(frame_range)
frame_lgv.grid(row=0, column=0, padx=1, pady=1)
optionLGV = tk.StringVar(value="LGV")
cb_radio = tk.Radiobutton(frame_lgv, text="CB", variable=optionLGV, value="CB")
cb_radio.grid(row=0, column=0, padx=1, pady=1, sticky='w')
lgv_radio = tk.Radiobutton(frame_lgv, text="LGV: ", variable=optionLGV, value="LGV")
lgv_radio.grid(row=0, column=1, padx=1, pady=1, sticky='w')

entry_lgv_range = tk.Entry(frame_range, fg="grey")
entry_lgv_range.grid(row=0, column=1, padx=5, pady=5)
create_placeholder(entry_lgv_range, "e.g., 1-5,11-17,20-25")

# Add a button to trigger table population
button_populate_table = tk.Button(root, text="Update Table", 
                                  bg="ghost white", 
                                  command=populate_table_from_inputs)
button_populate_table.grid(row=1, column=1, pady=10)
button_design(button_populate_table)

frame_ip = tk.Frame(root)
frame_ip.grid(row=3, column=0, padx=5, pady=5, sticky='e')

label_ip_help = tk.Label(frame_ip, text=" ? ", bd=2, relief='raised')
label_ip_help.grid(row=0, column=0, padx=5, pady=5)

ToolTip(label_ip_help, "The IP of the first element of the range")

label_ip = tk.Label(frame_ip, text="First IP: ")
label_ip.grid(row=0, column=1, padx=5, pady=5)

entry_base_ip = tk.Entry(frame_ip, fg="grey")
create_placeholder(entry_base_ip, "e.g., 172.20.3.10")
entry_base_ip.grid(row=0, column=2, padx=5, pady=5)
entry_base_ip.bind("<KeyRelease>", validate_base_ip)


# Button to delete the XML
delete_table_button = tk.Button(root, text="Delete Table", 
                                bg="ghost white", 
                                command=delete_whole_table)
delete_table_button.grid(row=3, column=1, pady=10)
button_design(delete_table_button)


frame_load = tk.Frame(root, bd=1, relief="groove")
frame_load.grid(row=4, column=0, columnspan=1, padx=15, pady=5, sticky='w')

load_label = tk.Label(frame_load, text="Load")
load_label.grid(row=0, column=0, padx=5, pady=0, sticky='w')
# Add a button to trigger the XML file selection and table population
load_xml_button = tk.Button(frame_load, text="StaticRoutes.xml", 
                            bg="ghost white", 
                            command=populate_table_from_xml)
load_xml_button.grid(row=1, column=0, padx=10, pady=5)
button_design(load_xml_button)

load_db3_button = tk.Button(frame_load, text="     Config.db3     ", 
                            bg="ghost white", 
                            command=populate_table_from_db3)
load_db3_button.grid(row=2, column=0, padx=10, pady=5)
button_design(load_db3_button)


frame_login = tk.Frame(root)
frame_login.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

frame_user = tk.Frame(frame_login)
frame_user.grid(row=0, column=0, padx=5, pady=0)

username_label = tk.Label(frame_user, text="Username:")
username_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')

username_entry = tk.Entry(frame_user, width=15)
username_entry.insert(0, "Administrator")
username_entry.grid(row=0, column=1, padx=5, pady=5)

frame_password = tk.Frame(frame_login)
frame_password.grid(row=1, column=0, padx=5, pady=0)

password_label = tk.Label(frame_password, text="Password:")
password_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')

password_entry = tk.Entry(frame_password, show="*", width=15)
password_entry.grid(row=0, column=1, padx=5, pady=5)

create_route_button = tk.Button(frame_login, text="Create Routes",
                                bg="ghost white",
                                command=create_tc_routes)
create_route_button.grid(row=0, column=2, padx=5, pady=5)
button_design(create_route_button)

# Add a frame to hold the Treeview and the scrollbar
frame_table = tk.Frame(root)
frame_table.grid(row=5, columnspan=3, padx=15, pady=10)

# Add a Treeview to display the data
treeview = ttk.Treeview(frame_table, columns=("Name", "Address", "NetId", "Type"), show="headings", height=10)
treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
setup_treeview()

# Add a vertical scrollbar to the Treeview
vsb = ttk.Scrollbar(frame_table, orient="vertical", command=treeview.yview)
vsb.pack(side=tk.RIGHT, fill=tk.Y)

# Configure the Treeview to use the scrollbar
treeview.configure(yscrollcommand=vsb.set)

# Define the column widths
treeview.column("Name", width=120)
treeview.column("Address", width=120)
treeview.column("NetId", width=120)
treeview.column("Type", width=50)

treeview.bind('<Delete>', delete_selected_record)
treeview.bind('<Double-1>', on_double_click)


frame_save_file = tk.Frame(root, bd=1, relief="groove")
frame_save_file.grid(row=6, column=0, columnspan=3, padx=5, pady=5)
save_label = tk.Label(frame_save_file, text="Save")
save_label.grid(row=0, column=0, padx=5, pady=0, sticky='w')
# Button to save the StaticRoutes.xml file
save_xml_button = tk.Button(frame_save_file, text="StaticRoutes.xml", 
                        bg="ghost white", 
                        command=save_routes_xml)
save_xml_button.grid(row=1, column=0, padx=5, pady=5)
button_design(save_xml_button)

# Button to save the ControlCenter.xml file
save_cc_button = tk.Button(frame_save_file, text="ControlCenter file", 
                             bg="ghost white", 
                             command=save_cc_xml)
save_cc_button.grid(row=1, column=2, padx=5, pady=5)
button_design(save_cc_button)

# Button to save WinSCP.ini file
save_winscp_button = tk.Button(frame_save_file, text="      WinSCP.ini      ", 
                             bg="ghost white", 
                             command=save_winscp_ini)
save_winscp_button.grid(row=1, column=3, padx=5, pady=5)
button_design(save_winscp_button)


# Create the context menu
context_menu = tk.Menu(treeview, tearoff=0)
context_menu.add_command(label="Delete", command=delete_selected_record_from_menu)

# Bind right-click to show the context menu
treeview.bind("<Button-3>", show_context_menu)

root.mainloop()

# leer config.db3 y llenar tabla con eso - DONE

# agregar rutas de ads

# add local route by default

# DELETE
# [Configuration\LastFingerprints]
# 172.20.2.68=20022:ssh=ecdsa-sha2-nistp384%20384%20iXnY+SMyoQRSUxJMzgWWA+yadddMZqqgM4dLPp/uHhs
# 172.20.2.68:20022:ssh=ecdsa-sha2-nistp384%20384%20iXnY+SMyoQRSUxJMzgWWA+yadddMZqqgM4dLPp/uHhs