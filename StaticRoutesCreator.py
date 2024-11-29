import sys
import os
import time
import platform
import threading
import re
import tkinter as tk
from tkinter import messagebox, filedialog, ttk, font
import xml.etree.ElementTree as ET
from xml.dom import minidom
import sqlite3
import winreg as reg
import configparser
import pyads
import socket
import select
import asyncio
import struct
import winreg
import paramiko
from threading import Thread
import logging

__version__ = '3.4.5'

default_file_path = os.path.join(r'C:\TwinCAT\3.1\Target', 'StaticRoutes.xml')

class ToolTip:
    def __init__(self, widget, text, delay=400, fade_duration=500):
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

        style = ttk.Style()
        style.configure("Tooltip.TLabel", background="white", relief='solid', borderwidth=1, font=("helvetica", "8", "normal"))

        label = ttk.Label(tw, text=self.text, 
                          style="Tooltip.TLabel"
                        #   justify='left',
                        #   background="white", relief='solid', borderwidth=1,
                        #   font=("helvetica", "8", "normal")
                          )
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

    xmlstr = minidom.parseString(ET.tostring(config)).toprettyxml(indent="    ")
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
    pretty_xml = reparsed.toprettyxml(indent="    ")
    
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

################################################ Validate inputs ##################################################################
good_input_bg = 'white'
bad_input_bg = '#fbcbcb' # light red
good_input_fg = 'black'
bad_input_fg = '#de021a'
placeholder_fg = 'grey'

def validate_entry(entry, style_name, validate_func):
    def inner_validate(*args):
        if entry.get() != placeholders[entry]:
            result = validate_func(entry)
            if result:
                style.configure(style_name, background=good_input_bg, foreground=good_input_fg)
            else:
                style.configure(style_name, background=bad_input_bg, foreground=bad_input_fg)
        else:
             style.configure(style_name, background=good_input_bg, foreground=placeholder_fg)
    return inner_validate

# Real-time validation functions
def validate_project(*args):
    project = entry_project.get().strip()
    if (project.isdigit() and len(project) == 4):
        return True
    elif not project or project == placeholders[entry_project]:
        return  False
    else:
        return False

def validate_range(*args):
    pattern = r"^\d+(-\d+)?(,\d+(-\d+)?)*$"
    input_range = entry_lgv_range.get().strip()
    
    # If the input is empty, reset to good input colors and return False
    if not input_range or input_range == placeholders[entry_lgv_range]:
        return False
    
    # Check if the input matches the pattern
    if re.match(pattern, input_range):
        ranges = input_range.split(',')
        
        # Check that each range is in increasing order, starts with a number greater than 0,
        # and does not have leading zeros
        for r in ranges:
            if '-' in r:
                start, end = r.split('-')
                if not start.isdigit() or not end.isdigit() or int(start) <= 0 or int(start) > int(end) or start != str(int(start)) or end != str(int(end)):
                    return False
            else:
                # If it's a single number, ensure it's greater than 0 and does not have leading zeros
                if not r.isdigit() or int(r) <= 0 or r != str(int(r)):
                    return False
        return True
    else:
        return False

def validate_base_ip(*args):
    base_ip = entry_base_ip.get().strip()
    if validate_ip(base_ip):
        return True
    elif not base_ip or base_ip == placeholders[entry_base_ip]:      
        return False
    else:
        return False

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

######################################## placeholders #######################################33
placeholders = {}

def create_placeholder(entry, placeholder_text, entry_style, placeholder_style):
    entry.insert(0, placeholder_text)
    entry.config(style=placeholder_style)
    entry.bind("<FocusIn>", lambda event: on_focus_in(entry, placeholder_text, entry_style))
    entry.bind("<FocusOut>", lambda event: on_focus_out(entry, placeholder_text, placeholder_style))
    placeholders[entry] = placeholder_text

def on_focus_in(entry, placeholder_text, entry_style):
    if entry.get() == placeholder_text:
        entry.delete(0, tk.END)
        entry.config(style=entry_style)

def on_focus_out(entry, placeholder_text, placeholder_style):
    if not entry.get():
        entry.insert(0, placeholder_text)
        entry.config(style=placeholder_style)

# Function to validate the inputs and create XML
def validate_and_create_xml():
    try:
        project = str(entry_project.get())
        if len(project) != 4:
            raise ValueError("Project number must be a 4 digit number")
    except ValueError as e:
        messagebox.showerror("Invalid input", str(e))
        return
    # if not validate_project():

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

def populate_table_from_xml(path=None):
    if not path:
        # Ask the user to select an XML file
        file_path = filedialog.askopenfilename(title="Select StaticRoutes file", 
                                            initialdir="C:\\TwinCAT\\3.1\\Target",
                                            filetypes=[("XML files", "*.xml")])
    else:
        file_path = path

    if file_path and not os.path.exists(file_path):
        print(f"The file {path} does not exist.")
        return
    
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
        for i in routes_table.get_children():
            routes_table.delete(i)
            
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
            routes_table.insert("", "end", values=item)
        # messagebox.showinfo("Success", "Data loaded successfully from the XML file.")

    update_tunnel_button_status()

############################# Populate table based on inputs ##############################

def populate_table_from_inputs():
    project = entry_project.get()
    lgv_range = entry_lgv_range.get()
    base_ip = entry_base_ip.get()
    is_tc3 = optionTC.get() == "TC3"
    # is_lgv = optionLGV.get() == "LGV"

    if not validate_project():
        messagebox.showerror("Invalid input", "Project number must be a 4 digit number")
        return 
    
    if not validate_range():
        messagebox.showerror("Invalid input", "Please enter a valid range")
        return
    
    if not validate_base_ip():
        messagebox.showerror("Invalid input", "Please enter a valid base IP address in the format 'xxx.xxx.xxx.xxx'")
        return
    
    lgvs = parse_range(lgv_range)

    # Parse the IPs based on the given base IP and LGV list
    ip_list = parse_ip(base_ip, lgvs)

    # Clear existing table data
    # for i in routes_table.get_children():
    #     routes_table.delete(i)

    # Loop through the parsed IPs and add to the table
    for i, current_ip in enumerate(ip_list):
        net_id = f"{current_ip}.1.1"
        route_name = f"CC{project}_{optionLGV.get()}{str(lgvs[i]).zfill(2)}"
        
        # Check if the record already exists in the table, ignoring the TC type and name
        record_exists = False
        for row in routes_table.get_children():
            existing_values = routes_table.item(row)["values"]
            if (existing_values[1] == current_ip and 
                existing_values[2] == net_id):
                record_exists = True
                # messagebox.showwarning("Duplicate Entry", f"The IP {current_ip} already exists.")
                break

        # Only add the record if it doesn't already exist
        if not record_exists:
            routes_table.insert("", "end", values=(
                route_name,
                current_ip,
                net_id,
                "TC3" if is_tc3 else "TC2"
            ))
    
    update_tunnel_button_status()

##################### Function to check for duplicates in the Treeview ######################
#Check for duplicates when input is the entry fields
# def is_duplicate(name, address, netid, type_tc):
#     for item in routes_table.get_children():
#         existing_values = routes_table.item(item, 'values')
#         if (name, address, netid, type_tc) == existing_values:
#             return True
#     return False

#Check for duplicates when input is the table directly
def is_duplicate(col_index, new_value, current_row):
    # Check for duplicates in the column except for the current editing row
    for item in routes_table.get_children():
        if item != current_row:
            if routes_table.item(item, 'values')[col_index] == new_value:
                return True
    return False

############################### Get IPs ###################################################
def parse_ip(base_ip, lgv_list):
    base_ip_prefix, start_ip = base_ip.rsplit('.', 1)
    start_ip = int(start_ip) - lgv_list[0]
    ip_list = []
    for lgv in lgv_list:
        if start_ip + lgv > 255: # IP digit must not be grater than 255
            messagebox.showinfo("Attention", f"Resulting IPs after {current_ip} out of range")
            return ip_list
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
    selected_item = routes_table.selection()
    if selected_item:
        routes_table.delete(selected_item)
    else:
        messagebox.showwarning("Selection Error", "Please select a record to delete.")

# With right click
def delete_selected_record_from_menu():
    selected_item = routes_table.selection()
    if selected_item:
        routes_table.delete(selected_item)

# Function to show the context menu
def show_context_menu(event):
    # Check if a record is selected
    selected_item = routes_table.identify_row(event.y)
    if selected_item:
        routes_table.selection_set(selected_item)
        context_menu.post(event.x_root, event.y_root)

# With DEL key
def delete_selected_record(event):
    selected_items = routes_table.selection()
    for item in selected_items:
        if item:
            routes_table.delete(item)

################################### Delete whole table ########################################

def delete_whole_table():
    for i in routes_table.get_children():
        routes_table.delete(i)

################################## Create StaticRoutes.xml from table ##########################
    
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
    xmlstr = minidom.parseString(ET.tostring(config, 'utf-8')).toprettyxml(indent="    ")

    # Write to a file
    with open(file_path, "w", encoding='utf-8') as f:
        f.write(xmlstr)

    messagebox.showinfo("Success", "StaticRoutes file has been created successfully. \nRemember to restart TwinCAT!!")

def save_routes_xml():
    if not get_table_data():
        messagebox.showerror("Attention", "Routes table is empty!")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xml",
                                             initialdir="C:\\TwinCAT\\3.1\\Target",
                                             initialfile="StaticRoutes.xml",
                                             filetypes=[("XML files", "*.xml")])
    if file_path:
        create_routes_xml_from_table(file_path)


def save_routes():
    if only_tc2_installed:
        save_routes_registry()
    else:
        save_routes_xml()

def save_routes_registry():
    data = get_table_data()
    i = 0
    total = len(data)
    for item in data:
        i+=1
        if not save_route_tc2(item):
            break

    if i==total:
        messagebox.showinfo("Success", "All routes added to the registry. \nRemember to restart TwinCAT!!")
    else:
        messagebox.showerror("Attention", "Unable to add some routes to the registry")
    

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
    xmlstr = minidom.parseString(ET.tostring(fleet, 'utf-8')).toprettyxml(indent="    ")

    # Write to a file
    with open(file_path, "w", encoding='utf-8') as f:
        f.write(xmlstr)

    messagebox.showinfo("Success", "ControlCenter file has been created successfully!")

def save_cc_xml():
    if not get_table_data():
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
    region = routes_table.identify("region", event.x, event.y)
    if region == "cell":
        column = routes_table.identify_column(event.x)
        row = routes_table.identify_row(event.y)
        col_index = int(column.replace("#", "")) - 1
        current_value = routes_table.item(row, 'values')[col_index]

        if col_index == 3:  # Assuming column 3 is the Type column
            create_combobox_for_type(column, row)
        else:
            create_entry_for_editing(column, row, col_index, current_value)


def create_combobox_for_type(column, row):
    bbox = routes_table.bbox(row, column)
    if not bbox:
        return
    
    combo_edit = ttk.Combobox(routes_table, values=["TC2", "TC3"], state="readonly")
    x, y, width, height = routes_table.bbox(row, column)
    combo_edit.place(x=x, y=y, width=width, height=height)

    def on_select(event):
        if combo_edit.winfo_exists():
            routes_table.set(row, column=column, value=combo_edit.get())
            combo_edit.destroy()
    
    def check_focus(event):
        # Destroy the Combobox if it is not the focus
        if event.widget != combo_edit:
            combo_edit.destroy()

    combo_edit.bind("<<ComboboxSelected>>", on_select)
    root.bind("<Button-1>", check_focus, add="+")  # Use "+" to add to existing bindings
    # combo_edit.focus()


def create_entry_for_editing(column, row, col_index, current_value):
    bbox = routes_table.bbox(row, column)
    if not bbox:
        return
    
    entry_edit = tk.Entry(routes_table, border=0)
    entry_edit.insert(0, current_value)
    x, y, width, height = routes_table.bbox(row, column)
    entry_edit.place(x=x, y=y, width=width, height=height)
    entry_edit.focus()
    entry_edit.select_range(0, tk.END)

    def save_edit(event):
        if entry_edit.winfo_exists():
            new_value = entry_edit.get()
            if is_duplicate(col_index, new_value, row):
                messagebox.showerror("Invalid Input", f"Duplicate value found for {routes_table.heading(col_index, 'text')}.")
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
            
            routes_table.set(row, column=column, value=new_value)

    def cancel_edit(event=None):
        if entry_edit.winfo_exists():
            entry_edit.destroy()

    entry_edit.bind("<Return>", save_edit)
    entry_edit.bind("<Escape>", lambda e: cancel_edit())
    entry_edit.bind("<FocusOut>", lambda e: cancel_edit())


################################## Sorting ################################################
def setup_routes_table():
    # Initialize the headings with custom names
    headings = {
        'Name': 'Route Name',
        'Address': 'IP Address',
        'NetId': 'AMS Net Id',
        'Type': 'Type'
    }
    
    for col in routes_table['columns']:
        routes_table.heading(col, text=headings[col], command=lambda _col=col: routes_table_sort_column(routes_table, _col, False), anchor='w')

def routes_table_sort_column(tv, col, reverse):
    # Retrieve all data from the routes_table
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
        tv.heading(column, text=heading_text, command=lambda _col=column: routes_table_sort_column(tv, _col, not reverse))

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
            return None

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
    if not validate_project():
        messagebox.showinfo("Attention", "Add project number first!")
        return
    
    db3_path = filedialog.askopenfilename(title="Select config.db3 file", 
                                          initialdir="C:\\Program Files (x86)\\Elettric80",
                                          filetypes=[("DB3 files", "*.db3")])
    if not db3_path:
        return
    
    table_agvs = "tbl_AGVs"
    rows_agvs = read_db3_file(db3_path, table_agvs)
    if rows_agvs is None:
        return
    
    table_param = "tbl_Parameter"
    rows_param = read_db3_file(db3_path, table_param)
    if rows_param is None:
        return
    
    # # print(columns, rows)
    
    # Clear the existing table data
    for i in routes_table.get_children():
        routes_table.delete(i)
    
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
        routes_table.insert("", "end", values=item)

################################# Split project and LGV number ######################################
def split_string(input_string):
    # Regular expression pattern for CCxxxxLGVxx or CCxxxx_LGVxx
    pattern_with_underscore = r"^CC\d{4}_(LGV|CB|BC)\d{2,3}$"
    pattern_without_underscore = r"^CC\d{4}(LGV|CB|BC)\d{2,3}$"
    
    # Check if the input string matches the expected patterns
    if re.match(pattern_with_underscore, input_string):
        # Split by underscore if present
        parts = input_string.split('_')
    elif re.match(pattern_without_underscore, input_string):
         # Use regular expression to split between the numeric and alphanumeric parts
        match = re.match(r"(CC\d{4})(LGV\d{2}|CB\d{2}|BC\d{2,3})", input_string)
        parts = [match.group(1), match.group(2)]
    else:
        # Raise a ValueError if the string does not match the expected format
        messagebox.showerror("Error", f"Route name '{input_string}' does not match the expected format: CCxxxxLGV/CB/BCxx or CCxxxx_LGV/CB/BCxx")
        parts = None
    return parts


def parse_route_name(input_name):
    # Updated regex to capture the section correctly and ensure the last part is treated as the name
    pattern = r"^(?P<section>CC\d+(?:_[\w\-]*)?)_(?P<name>[A-Za-z]*\d{1,3})$"

    match = re.match(pattern, input_name)
    if match:
        section = match.group("section")
        name = match.group("name")

        # Error: Missing section or invalid format (e.g., 'LGV02')
        if not section or not section.startswith("CC"):
            messagebox.showerror(
                "Invalid Input", f"Missing or invalid section in '{input_name}'. Please provide a valid section starting with 'CC'."
            )
            return None, None

        # Handle names that are just numbers by defaulting to 'LGV'
        if name.isdigit():
            name = f"LGV{name.zfill(2)}"
        else:
            # Pad single-digit names with leading zeros if necessary
            if re.search(r'\d$', name) and len(re.search(r'\d+$', name).group()) == 1:
                name = name[:-1] + f"0{name[-1]}"
        
        return section, name
    else:
        messagebox.showerror("Invalid Input", f"'{input_name}' is not in a valid format.")
        return None, None
    
################################### Create ini file for WinSCP connections ##########################
# Function to set the custom INI path in the Windows Registry
def set_custom_ini_path(ini_path):
    key_path = r'Software\Martin Prikryl\WinSCP 2\Configuration'

    try:
        # Open the key for reading
        key = reg.OpenKey(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_READ)
        try:
            config_storage, _ = reg.QueryValueEx(key, "ConfigurationStorage")
            custom_ini_file, _ = reg.QueryValueEx(key, "CustomIniFile")

            # Check if the values are already set correctly
            if config_storage == 1 and custom_ini_file == ini_path:
                print("Registry keys are already set correctly.")
                reg.CloseKey(key)
                return True
        except FileNotFoundError:
            # Values not set, proceed to create/update them
            pass

        reg.CloseKey(key)

        # Open the key for writing
        key = reg.OpenKey(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_SET_VALUE)
        reg.SetValueEx(key, "ConfigurationStorage", 0, reg.REG_DWORD, 1)
        reg.SetValueEx(key, "CustomIniFile", 0, reg.REG_SZ, ini_path)
        reg.CloseKey(key)

        print("Registry keys updated successfully.")
        return True

    except Exception as e:
        print(f"Failed to set registry key: {e}")
        messagebox.showerror("Error", f"Failed to set path {ini_path}, run app as administrator")
        return False
    
# Function to check if a HostName already exists in the INI file
def hostname_exists(config, host_name):
    for section in config.sections():
        if config.has_option(section, 'HostName') and config.get(section, 'HostName') == host_name:
            return True
    return False

# Function to create a session in the winscp.ini file
def create_winscp_ini_from_table(ini_path, data):
    # Ensure the directory for the INI file exists
    ini_dir = os.path.dirname(ini_path)
    os.makedirs(ini_dir, exist_ok=True)

    # Create config parser and read the INI file (if it exists)
    config = configparser.ConfigParser()
    try:
        if os.path.exists(ini_path):
            config.read(ini_path)

    except configparser.DuplicateOptionError as e:
        print(f"Duplicate option found and skipped: {e}")

    except configparser.ParsingError as e:
        print(f"Error parsing INI file: {e}")
        return "Failed to read the INI file due to a parsing error."

    # data = get_table_data()
    repeated = 0
    total = 0
    for row in data:
        total += 1 
        name, address, netid, tc_type = row
        name_parts = parse_route_name(str(name))
        if name_parts is None:
            return
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

    try:
        with open(ini_path, 'w') as configfile:
            config.write(configfile)
    except Exception as e:
        print(f"An error occurred while writing the INI file: {e}") 
        messagebox.showerror("Error", f"An error occurred while writing the INI file: {e}")
    
    # Set the custom INI path in the registry
    if not set_custom_ini_path(ini_path):
        return "Failed to set the custom INI path in the registry."

    messagebox.showinfo("Success", f"Session created successfully in {ini_path} with {repeated} repeated routes out of {total}")

def save_winscp_ini():
    # Custom path for the INI file (not in Roaming)
    file_path = r'C:\WinSCPConfig\WinSCP.ini'  # Adjust this path if needed
    data = get_table_data()
    if not data:
        messagebox.showerror("Attention", "Routes table is empty!")
        return
    # file_path = filedialog.asksaveasfilename(defaultextension=".ini",
    #                                          initialdir= os.path.join(os.path.expanduser("~"), "Documents"),
    #                                          initialfile="WinSCP.ini",
    #                                          filetypes=[("INI files", "*.ini")])
    if file_path:
        create_winscp_ini_from_table(file_path, data)

############################################################## RDP connection #################################################################
def open_rdp_connection():
    print("Set RDP connection")

############################################################## SSH tunneling config #################################################################
SSH_CONFIG_FILE = "ssh_config.xml"

# Global variable to track active SSH client
active_ssh_client = None

default_tunnel_data = [
            ("40101", "192.168.11.61", "2122", "PLS Front ETH"),
            ("40102", "192.168.11.62", "2122", "PLS Rear ETH"),
            ("40105", "192.168.11.65", "2122", "PLS Lateral Left ETH"),
            ("40106", "192.168.11.66", "2122", "PLS Lateral Right ETH"),
            ("5900",  "192.168.11.6",  "5900", "Exor OnBoard VNC")
        ]

def update_tunnel_button_status():
    """
    Enable the tunnel setup button if at least one element in the table is TC3.
    Disable it otherwise.
    """
    # Get all items in the static routes table
    all_items = routes_table.get_children()
    has_tc3 = any(routes_table.item(item)["values"][3] == "TC3" for item in all_items)  # Assuming "Type" is in the 4th column
    
    if has_tc3:
        setup_tunnel_button.config(state=tk.NORMAL)
    else:
        setup_tunnel_button.config(state=tk.DISABLED)

    # Call this method every time the static routes table is updated
    # Example: after loading new data or modifying entries
    # update_tunnel_button_status()

def update_ssh_menu_status():
    """
    Enable SSH option only if selected element is TC3
    """
    selection = routes_table.selection()
    if selection:
        if routes_table.item(selection)["values"][3] == "TC3":
            context_menu.entryconfig("SSH Tunnel", state="normal")
        else:
            context_menu.entryconfig("SSH Tunnel", state="disabled")

def update_ssh_state(*args):
    update_tunnel_button_status()
    update_ssh_menu_status()

logging.basicConfig(level=logging.DEBUG)

tunnel_connection_in_progress = False
def create_ssh_tunnel():
    """Create SSH tunnels for the selected LGV."""
    global active_ssh_client, tunnel_connection_in_progress

    # Check if a tunnel is already active
    if active_ssh_client and active_ssh_client.get_transport() and active_ssh_client.get_transport().is_active():
        messagebox.showwarning("Tunnel Active", "An SSH tunnel is already active. Close it before creating a new one.")
        return

    tunnels = get_tunnels()

    ssh_username = username_entry.get()
    ssh_password = password_entry.get()

    selected_item = routes_table.selection()

    if not tunnels:
        messagebox.showinfo("No Tunnels", "No tunnels found to create.")
        return

    if not ssh_username:
        messagebox.showwarning("Attention", "Input username")
        print("Input user")
        return
    if not ssh_password:
        messagebox.showwarning("Attention", "Input password")
        print("Input password")
        return
    
    if tunnel_connection_in_progress:
        print("Tunnel connection in progress")
        return

    lgv = routes_table.item(selected_item)["values"][0]
    print(f"SSH Tunnel created for {lgv}")

    ssh_host = routes_table.item(selected_item)["values"][1]

    tunnel_connection_in_progress = True

    def establish_tunnels():
        """Establish SSH tunnels."""
        global active_ssh_client, tunnel_connection_in_progress
        try:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(ssh_host, port=20022, username=ssh_username, password=ssh_password, timeout=5)

            active_ssh_client = client  # Set the global active client

            transport = client.get_transport()
            transport.set_keepalive(30)  # Send a keepalive packet every 30 seconds
            active_tunnels = []

            for tunnel in tunnels:
                local_port = int(tunnel["Local Port"])
                remote_ip = tunnel["Remote IP"]
                remote_port = int(tunnel["Remote Port"])

                try:
                    # Establish forwarding
                    channel = transport.open_channel(
                        "direct-tcpip",
                        (remote_ip, remote_port),  # Remote target
                        ("127.0.0.1", local_port)  # Local source
                    )
                    active_tunnels.append(f"{local_port} -> {remote_ip}:{remote_port}")
                    print(f"Tunnel created: {local_port} -> {remote_ip}:{remote_port}")
                except Exception as tunnel_error:
                    print(f"Error creating tunnel {local_port} -> {remote_ip}:{remote_port}: {tunnel_error}")

            if not active_tunnels:
                raise Exception("No tunnels could be established.")

            # Open confirmation window
            show_tunnel_window(client, active_tunnels, lgv)

        except paramiko.AuthenticationException:
            messagebox.showerror("Authentication Error", "Invalid username or password for SSH.")
        except paramiko.SSHException as e:
            messagebox.showerror("SSH Error", f"SSH connection to {lgv} failed: {e}")
        except Exception as e:
            print(f"Error creating tunnels for {lgv}: {e}")
            messagebox.showerror("Error", f"Failed to create SSH tunnels for {lgv}: {e}")
            if active_ssh_client:
                try:
                    active_ssh_client.close()
                    print("Closed active SSH client due to errors.")
                except Exception as cleanup_error:
                    print(f"Error during cleanup: {cleanup_error}")
                active_ssh_client = None
        finally:
            tunnel_connection_in_progress = False


    # Run tunnel creation in a thread to avoid blocking the UI
    tunnel_thread = Thread(target=establish_tunnels, daemon=True)
    tunnel_thread.start()

def show_tunnel_window(client, active_tunnels, host):
    """Display a window showing active tunnels and allow closing them."""
    global active_ssh_client

    tunnel_window = tk.Toplevel(root)
    tunnel_window.title("Active Tunnels")
    
    window_width = 400
    window_lenght = 300
    tunnel_window.geometry(f"{window_width}x{window_lenght}")
    tunnel_window.minsize(window_width, window_lenght)

    # Show LGV information
    tk.Label(tunnel_window, text=f"Tunnels are active for {host}:", font=("Arial", 14)).pack(pady=10)
    
    # Show tunnel status
    tunnel_list = tk.Text(tunnel_window, wrap="word", height=10, width=40)
    tunnel_list.pack(pady=10)

    for tunnel in active_tunnels:
        tunnel_list.insert("end", f"{tunnel}\n")
    tunnel_list.configure(state="disabled")

    def close_tunnels():
        """Close all active tunnels and destroy the window."""
        global active_ssh_client
        try:
           if client and client.get_transport() and client.get_transport().is_active():
                client.close()
                print(f"All tunnels for {host} have been closed.")
                messagebox.showinfo("Tunnels Closed", f"All tunnels for {host} have been closed.")
                active_ssh_client = None  # Reset active client
        except Exception as e:
            messagebox.showerror("Error", f"Failed to close tunnels: {e}")
        finally:
            tunnel_window.destroy()
    
    # Bind the tunnel window close event to cleanup
    tunnel_window.protocol("WM_DELETE_WINDOW", close_tunnels)

    close_button = ttk.Button(tunnel_window, text="Close Tunnels", command=close_tunnels)
    close_button.pack(pady=10)

def get_tunnels():
    """Load tunnels from XML file or use default ones"""
    global default_tunnel_data

    tunnels = []

    if os.path.exists(SSH_CONFIG_FILE):
        try:
            tree = ET.parse(SSH_CONFIG_FILE)
            root = tree.getroot()

            for tunnel in root.findall("Tunnel"):
                tunnels.append({
                    "Local Port" : tunnel.find("LocalPort").text,
                    "Remote IP"  : tunnel.find("RemoteIP").text,
                    "Remote Port": tunnel.find("RemotePort").text,
                })
        except Exception as e:
            print(f"Error reading XML file {e}")
    
    if not tunnels:
        tunnels = [
            {"Local Port" : row[0], "Remote IP"  : row[1], "Remote Port": row[2]}
            for row in default_tunnel_data
        ]
    
    return tunnels


"""
def on_main_app_close():
    Handle the main application close event.
    try:
        if client and client.get_transport() and client.get_transport().is_active():
            client.close()
            print("All tunnels closed as the main app is closing.")
    except Exception as e:
        print(f"Error closing tunnels during app exit: {e}")
    root.destroy()
"""
##################################################### Setup SSH window #############################################################################
ssh_config_window = None


def open_ssh_config_window_cond():
    global ssh_config_window

    if ssh_config_window is not None and ssh_config_window.winfo_exists():
        ssh_config_window.lift()
        ssh_config_window.focus_force()
    else:
        open_ssh_config_window()

def open_ssh_config_window():
    global ssh_config_window, is_saved

    # Track if the table data has been saved
    is_saved = False

    ssh_config_window = tk.Toplevel(root)
    ssh_config_window.title("Setup SSH")

    window_width = 500
    window_lenght = 300
    ssh_config_window.geometry(f"{window_width}x{window_lenght}")
    ssh_config_window.minsize(window_width, window_lenght)

    # Dictionary to maintain custom headings
    headings = {
        'Local Port' : 'Local Port',
        'Remote IP'  : 'Remote IP',
        'Remote Port': 'Remote Port',
        'Description': 'Description'
    }

    def setup_tunnel_table():
        for col in tunnel_table['columns']:
            tunnel_table.heading(col, text=headings[col], command=lambda _col=col: table_sort_column(tunnel_table, _col, False), anchor='w')

    def table_sort_column(tv, col, reverse):
        # Retrieve all data from the treeview
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Sort the data
        l.sort(reverse=reverse, key=lambda t: natural_keys(t[0]))

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Change the heading to show the sort direction
        for column in tv['columns']:
            heading_text = headings[column] + (' ↓' if reverse and column == col else ' ↑' if not reverse and column == col else '')
            tv.heading(column, text=heading_text, command=lambda _col=column: table_sort_column(tv, _col, not reverse))

    def natural_keys(text):
        """
        Alphanumeric (natural) sort to handle numbers within strings correctly
        """
        return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]

    def initialize_tunnel_table():
        """Initialize the tunnel table based on the presence of the XML file."""
        if os.path.exists(SSH_CONFIG_FILE):
            load_table_from_xml(SSH_CONFIG_FILE)
        else:
            init_tunnel_table()

    def init_tunnel_table():
        """Initialize the table with default values."""
        global default_tunnel_data

        for row in default_tunnel_data:
            add_row(tunnel_table, *row)

    
    def save_table_to_xml(filename=SSH_CONFIG_FILE):
        """Save table data to an XML file."""
        global is_saved
        # Create the root element
        root = ET.Element("Tunnels")

        # Add data from the table to the XML structure
        for row_id in tunnel_table.get_children():
            values = tunnel_table.item(row_id)["values"]
            tunnel = ET.SubElement(root, "Tunnel")
            ET.SubElement(tunnel, "LocalPort").text   = str(values[0])
            ET.SubElement(tunnel, "RemoteIP").text    = str(values[1])
            ET.SubElement(tunnel, "RemotePort").text  = str(values[2])
            ET.SubElement(tunnel, "Description").text = str(values[3])

        # Convert the XML structure to a string and prettify it
        xml_str = ET.tostring(root, encoding="unicode")
        pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="    ")

        # Write the pretty XML to the file
        full_path = os.path.abspath(filename)
        with open(full_path, "w") as file:
            file.write(pretty_xml)

        is_saved = True
        messagebox.showinfo("Save Successful", f"Table data saved to:\n{full_path}")

    def load_table_from_xml(filename=SSH_CONFIG_FILE):
        """Load table data from an XML file."""
        try:
            tree = ET.parse(filename)
            root = tree.getroot()

            for tunnel in root.findall("Tunnel"):
                local_port = tunnel.find("LocalPort").text
                remote_ip = tunnel.find("RemoteIP").text
                remote_port = tunnel.find("RemotePort").text
                description = tunnel.find("Description").text
                add_row(tunnel_table, local_port, remote_ip, remote_port, description)
        except FileNotFoundError:
            print(f"{filename} not found. Starting with an empty table.")

    def on_window_close():
        """Handle the window close event"""
        if not is_saved:
            result  = messagebox.askyesnocancel(
                "Unsaved Changes",
                "You have unsaved changes. Do you want to save them before closing?"
            )
            if result is None:
                return
            elif result:
                save_table_to_xml()
        ssh_config_window.destroy()

    ssh_config_window.protocol("WM_DELETE_WINDOW", on_window_close)

    def mark_as_unsaved(event=None):
        global is_saved
        is_saved = False
    

    # Input frame for adding data
    input_button_frame = tk.Frame(ssh_config_window)
    input_button_frame.pack(fill=tk.X, pady=5, padx=5)

    # Use grid for input fields and buttons
    ttk.Label(input_button_frame, text="Local Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    local_port_entry = ttk.Entry(input_button_frame, width=8)
    local_port_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(input_button_frame, text="Remote Port:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    remote_port_entry = ttk.Entry(input_button_frame, width=8)
    remote_port_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(input_button_frame, text="Remote IP:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
    remote_ip_entry = ttk.Entry(input_button_frame, width=25)
    remote_ip_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")

    ttk.Label(input_button_frame, text="Description:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
    description_entry = ttk.Entry(input_button_frame, width=25)
    description_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")

    def add_row_from_inputs():
        """Add a row to the table using data from the input fields."""
        local_port = local_port_entry.get()
        remote_ip = remote_ip_entry.get()
        remote_port = remote_port_entry.get()
        description = description_entry.get()

        if not (local_port and remote_ip and remote_port and description):
            messagebox.showwarning("Input Error", "All fields must be filled.")
            return

        add_row(tunnel_table, local_port, remote_ip, remote_port, description)

        # # Clear the input fields
        # local_port_entry.delete(0, tk.END)
        # remote_ip_entry.delete(0, tk.END)
        # remote_port_entry.delete(0, tk.END)
        # description_entry.delete(0, tk.END)

    add_button = ttk.Button(input_button_frame, text="Add Data", command=add_row_from_inputs)
    add_button.grid(rowspan=2, row=0, column=4, columnspan=4, pady=10, padx=10)

    add_button.bind("<Button-1>", mark_as_unsaved)

    table_frame = tk.Frame(ssh_config_window)
    table_frame.pack(fill=tk.Y, expand=True, pady=10)
    # Create the Treeview widget (tunnel_table)
    columns = ("Local Port", "Remote IP", "Remote Port", "Description")
    tunnel_table = ttk.Treeview(table_frame, columns=columns, show="headings", height=5)
    
    tunnel_table.column("Local Port", width=70, anchor='w')
    tunnel_table.column("Remote IP", width=100, anchor='w')
    tunnel_table.column("Remote Port", width=80, anchor='w')
    tunnel_table.column("Description", width=210, anchor='w')

    setup_tunnel_table()

    tunnel_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10,0))

    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tunnel_table.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    tunnel_table.configure(yscrollcommand=vsb.set)

    # Bind the DEL key to delete rows
    tunnel_table.bind("<Delete>", lambda event: delete_selected_row(tunnel_table))
    tunnel_table.bind("<Key>", mark_as_unsaved)

    save_table_button = ttk.Button(ssh_config_window, text="  Save Table  ", command=save_table_to_xml)
    save_table_button.pack(side=tk.TOP, fill=tk.Y, pady=(0,10), padx=10)

    def add_row(tunnel_table, local_port="", remote_ip="", remote_port="", description=""):
        """Add a row with specified data to the table."""
        tunnel_table.insert("", "end", values=(local_port, remote_ip, remote_port, description))

    def delete_selected_row(tunnel_table):
        """Delete the selected row(s) from the table."""
        selected_items = tunnel_table.selection()
        if selected_items:
            for item in selected_items:
                tunnel_table.delete(item)
        else:
            messagebox.showwarning("Delete Row", "No row selected to delete.")

    initialize_tunnel_table()


######################################################## Create TC Routes ############################################################################

def string_to_byte_format(ip_string):
    # Split the string by the dot '.'
    parts = ip_string.split('.')
    
    # Convert each part to an integer and then to a byte
    byte_representation = bytes(int(part) for part in parts)
    
    return byte_representation

def get_local_ams_netid():
    ams_net_id=None
    try:
        pyads.open_port()
        ams_net_id = pyads.get_local_address().netid
        print (ams_net_id)
    except Exception as e:
        print(f"Unexpected error: {e} \nCheck if TwinCAT on local machine is running")
    finally:
        pyads.close_port()

    return ams_net_id

class TcpStateObject:
    def __init__(self):
        self.data = bytearray(1024)  # Buffer size, adjust as needed
        self.CurrentIndex = 0

class RouteManager:
    def __init__(self):
        self.AddRouteSuccess = False
        self.AddRouteError = False
        self._remoteAMSNetID = None
        self.ADSErrorCode = 0
        self.UDPSocket = None
        self.RouteAdded = False

    async def EZRegisterToRemote(self, local_name, local_ip, my_ams_net_id, username, password, remote_ip, use_static_route):
        print(f"Starting EZRegisterToRemote for PLC {remote_ip}...")

        my_ip_address = local_ip
        router_table_name = "TCP_" + local_name
        int_send_length = 27 + len(router_table_name) + 15 + len(username) + 5 + len(password) + 5 + len(my_ip_address) + 1

        if not use_static_route:
            int_send_length += 8

        sendbuf = bytearray(int_send_length + 1)

        sendbuf[0] = 3
        sendbuf[1] = 102
        sendbuf[2] = 20
        sendbuf[3] = 113
        sendbuf[4] = 0
        sendbuf[5] = 0
        sendbuf[6] = 0
        sendbuf[7] = 0
        sendbuf[8] = 6
        sendbuf[9] = 0
        sendbuf[10] = 0
        sendbuf[11] = 0

        # Copy AMS Net ID into the buffer at the correct position
        sendbuf[12:18] = my_ams_net_id

        sendbuf[18] = 16
        sendbuf[19] = 39

        if use_static_route:
            sendbuf[20] = 5
        else:
            sendbuf[20] = 6

        sendbuf[21] = 0
        sendbuf[22] = 0
        sendbuf[23] = 0
        sendbuf[24] = 12
        sendbuf[25] = 0
        sendbuf[26] = len(router_table_name) + 1
        sendbuf[27] = 0

        i = 28
        sendbuf[i:i + len(router_table_name)] = router_table_name.encode('ascii')
        i += len(router_table_name)

        sendbuf[i] = 0
        i += 1

        sendbuf[i] = 7
        sendbuf[i + 1] = 0
        sendbuf[i + 2] = 6
        sendbuf[i + 3] = 0
        i += 4

        # Copy AMS Net ID again into the buffer
        sendbuf[i:i + 6] = my_ams_net_id
        i += 6

        sendbuf[i] = 13
        sendbuf[i + 1] = 0
        i += 2

        sendbuf[i] = len(username) + 1
        sendbuf[i + 1] = 0
        i += 2

        sendbuf[i:i + len(username)] = username.encode('ascii')
        i += len(username)

        sendbuf[i] = 0
        i += 1

        sendbuf[i] = 2
        sendbuf[i + 1] = 0

        sendbuf[i + 2] = len(password) + 1
        sendbuf[i + 3] = 0
        i += 4

        sendbuf[i:i + len(password)] = password.encode('ascii')
        i += len(password)

        sendbuf[i] = 0
        i += 1

        sendbuf[i] = 5
        sendbuf[i + 1] = 0
        sendbuf[i + 2] = len(my_ip_address) + 1
        sendbuf[i + 3] = 0
        i += 4

        sendbuf[i:i + len(my_ip_address)] = my_ip_address.encode('ascii')
        i += len(my_ip_address)

        sendbuf[i] = 0
        i += 1

        if len(sendbuf) >= i + 8:
            sendbuf[i:i + 8] = struct.pack('BBBBBBBB', 9, 0, 4, 0, 1, 0, 0, 0)

        print(f"Sending message to {remote_ip}: {sendbuf.hex()}")

        # Now create the UDP socket and send the message asynchronously
        address = (remote_ip, 48899)
        self.UDPSocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.UDPSocket.settimeout(5.0)
        self.UDPSocket.setblocking(False)
        self.UDPSocket.connect(address)

        try:
            retries = 0
            state = TcpStateObject()
            c = 0
            timeout_occurred = False
            while retries < 1 and not self.AddRouteSuccess and not self.AddRouteError:
                self.UDPSocket.send(sendbuf)
                print(f"Message sent to {remote_ip}. Polling for response... (Attempt {retries + 1})")

                while not self.AddRouteSuccess and not self.AddRouteError and c < 80:
                    readable, _, _ = select.select([self.UDPSocket], [], [], 2.0)
                    if readable:
                        await self.DataReceivedA(self.UDPSocket, state)
                    else:
                        print(f"Timeout while waiting for data from {remote_ip}")
                        timeout_occurred = True
                        break

                    c += 1

                retries += 1

            if self.AddRouteSuccess:
                self.RouteAdded = True
                print(f"Route added successfully for PLC at {remote_ip}")
            elif timeout_occurred:
                self.RouteAdded = False
                print(f"Route was not added for {remote_ip} due to timeout.")
                print(f"No response from remote system {remote_ip}. Make sure firewall is off and check username, password, and computer name.")
            elif self.AddRouteError:
                self.RouteAdded = False
                print("Error encountered while adding route.")
                print("Error setting up remote system, check TwinCATCom for username, password, and computer name.")
            else:
                self.RouteAdded = False
                print(f"Failed to add route for {remote_ip} due to select timeout.")

        finally:
            self.UDPSocket.close()
            print(f"Socket closed for {remote_ip}.")

    async def DataReceivedA(self, udp_socket, state_obj):
        try:
            # Receive UDP message, block call, wait for data 
            bytes_received = udp_socket.recv(len(state_obj.data) - state_obj.CurrentIndex)

            # Update buffer
            state_obj.CurrentIndex += len(bytes_received)
            state_obj.data[state_obj.CurrentIndex:state_obj.CurrentIndex + len(bytes_received)] = bytes_received

            if state_obj.CurrentIndex > 31:
                AMSNetID = f"{state_obj.data[12]}.{state_obj.data[13]}.{state_obj.data[14]}.{state_obj.data[15]}.{state_obj.data[16]}.{state_obj.data[17]}"
                self._remoteAMSNetID = AMSNetID
                self.ADSErrorCode = state_obj.data[28] + state_obj.data[29] * 256

                if state_obj.data[27] == 0 and state_obj.data[28] == 0 and state_obj.data[29] == 0 and state_obj.data[30] == 0:
                    self.AddRouteSuccess = True
                    print("SUCCESS!!!!")
                    print(f"Route added successfully. Remote AMSNetID: {self._remoteAMSNetID}")
                else:
                    self.AddRouteError = True
                    self._remoteAMSNetID = "(null AMSID)"
                    print("Route addition failed. Error in response.")
                udp_socket.close()

            else:
                # Continue receiving asynchronously until enough data is received
                await self.DataReceivedA(udp_socket, state_obj)

        except socket.timeout:
            print("Socket timed out waiting for a response")
        
        except BlockingIOError:
            print("No data available right now, try again later")

        except Exception as e:
            print(f"Error receiving data: {e}")
            return

def get_items_for_routes():
    items = []
    selected_items = routes_table.selection()

    if selected_items: 
        for item in selected_items:
            # Append the item reference only
            items.append(item)
    else:
        # If no selection, return all items
        items = routes_table.get_children()
    
    return items

# Define global variables
route_creation_lock = threading.Lock()  # Lock for route creation threads
active_route_creation_threads = 0       # Counter for active route creation threads
lock = threading.Lock()                 # Lock for testing routes
failed_routes = []

# Modified to create routes for red-tagged entries (failed connections)
def create_tc_routes():

    # This gets either the user selection or the whole table
    items = get_items_for_routes()
    print(items)
    red_items = [item for item in items if 'red' in routes_table.item(item, 'tags')]
    print(red_items)
    
    if not red_items:
        print("No failed connections to create routes for.")
        return

    username = username_entry.get()
    password = password_entry.get()
    
    if not username or not password:
        messagebox.showerror("Attention", "Username and password are required!")
        return
    
    local_ams_net_id = get_local_ams_netid()
    
    system_name = platform.node()

    # Clear previous failed routes log
    failed_routes.clear()

    with route_creation_lock:
        global active_route_creation_threads
        active_route_creation_threads = len(red_items)

    start_spinner(190, 133)

    for item in red_items:
        entry = routes_table.item(item)["values"]
        # Start the creation process in a new thread for each red-tagged entry
        threading.Thread(target=create_and_retest_route, args=(entry, username, password, local_ams_net_id, system_name)).start()

def create_and_retest_route(entry, username, password, local_ams_net_id, system_name):
    name, remote_ip, ams_net_id, type_ = entry

    local_net_id = local_ams_net_id.split('.')
    local_netid_ip = '.'.join(local_net_id[:4])
    ams_net_id_bit = string_to_byte_format(local_ams_net_id)

    port = 851 if type_ == 'TC3' else 801

    try:
        # Create route using the route manager
        route_manager = RouteManager()
        asyncio.run(route_manager.EZRegisterToRemote(system_name, local_netid_ip, ams_net_id_bit, username, password, remote_ip, use_static_route=True))

        if route_manager.AddRouteSuccess:
            # After route creation, retest the connection
            connection_ok = test_connection(ams_net_id, port, name)

            # Update the UI with the result
            routes_table.after(0, lambda: update_ui_with_result_retest(name, connection_ok))
        else:
            # Log is route was not successfully added
            raise Exception(f"Route was not added for {remote_ip}")

    except Exception as e:
        print(f"Error during route creation or testing for {name}: {e}")
        # Add the failed route to the failed_routes list
        with route_creation_lock:
            failed_routes.append(f"{name} - {remote_ip} (Reason: {e})")
    
    finally:
        # Decrement the active route creation threads counter
        with route_creation_lock:
            global active_route_creation_threads
            active_route_creation_threads -= 1
            if active_route_creation_threads == 0:
                routes_table.after(0, lambda: stop_spinner())
                routes_table.after(0, lambda: routes_table.selection_remove(routes_table.selection()))
                routes_table.after(0, lambda: log_failed_routes())  # Log the failed routes

def log_failed_routes():
    if failed_routes:
        print("The following routes failed to be created:")
        for route in failed_routes:
            print(route)
    else:
        print("All routes were created successfully.")

# Function to update the UI with the result
def update_ui_with_result_retest(name, connection_ok):
    for item in routes_table.get_children():
        if routes_table.item(item, 'values')[0] == name:
            color = 'green' if connection_ok else 'red'
            routes_table.item(item, tags=(color,))
            break


def save_route_tc2(entry, flags=0, timeout=0, transport_type=1):
    route_name, address, net_id, _ = entry
    try:
        reg_path = r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT\Remote"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_SET_VALUE)

        # Create a new subkey for the route using route_name (e.g., "CC1393_01")
        route_key = winreg.CreateKey(key, route_name)
        
        # Set the route details
        winreg.SetValueEx(route_key, "Address", 0, winreg.REG_SZ, address)

        # Convert the AMS Net ID from string (e.g., "172.168.11.101.1.1") to binary format
        net_id_bytes = struct.pack('6B', *[int(x) for x in net_id.split('.')])
        winreg.SetValueEx(route_key, "AmsNetId", 0, winreg.REG_BINARY, net_id_bytes)

        # Set other values (Flags, Timeout, TransportType)
        winreg.SetValueEx(route_key, "Flags", 0, winreg.REG_DWORD, flags)
        winreg.SetValueEx(route_key, "Timeout", 0, winreg.REG_DWORD, timeout)
        winreg.SetValueEx(route_key, "TransportType", 0, winreg.REG_DWORD, transport_type)
        
        winreg.CloseKey(route_key)
        winreg.CloseKey(key)
        print(f"Route {route_name} added to TwinCAT2 registry.")
        return True
    except PermissionError:
        print("Access denied: Please run the application as administrator.")
        messagebox.showerror("Error", "Access denied: Please run the application as administrator.")
        return False
    except Exception as e:
        print(f"Failed to save route: {e}")
        return False


only_tc2_installed = False

def check_twinCAT_version():
    global only_tc2_installed
    try:
        # Check if TwinCAT3 is installed
        tc3_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT3")
        print("TwinCAT3 is installed.")
        # messagebox.showinfo("Attention", "TwinCAT3 is installed.")
        winreg.CloseKey(tc3_key)
        return "TC3"
    except FileNotFoundError:
        pass
    
    try:
        # Check if TwinCAT2 is installed
        tc2_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT")
        print("TwinCAT2 is installed.")
        winreg.CloseKey(tc2_key)
        only_tc2_installed = True
        messagebox.showinfo("Attention", "Only TwinCAT2 is installed, make sure app is running as administrator to save routes properly")
        return "TC2"
    except FileNotFoundError:
        print("Neither TwinCAT3 nor TwinCAT2 are installed.")
        return None

    
############################################### Test Routes ##################################################################
# Not used
def test_tc_routes_no_thread():
    # start_spinner(210, 225)
    data = get_data_for_routes()
    if not data:
        messagebox.showerror("Attention", "Routes table is empty!")
        stop_spinner()
        return None
    
    for entry in data:
        test_route_and_update_ui(entry,)


# Track the number of active threads
max_threads = 5
semaphore = threading.Semaphore(max_threads)
active_threads = 0
lock = threading.Lock()

def test_tc_routes():
    global active_threads
    if active_threads != 0:
        return
    
    data = get_data_for_routes()
    if not data:
        messagebox.showerror("Attention", "Routes table is empty!")
        stop_spinner()
        return None
    
    # Set back to black when testing again
    # only tested routes if selected, if not all of them
    # for item in routes_table.get_children():
    selected = routes_table.selection()
    if selected:
        for item in selected:
            routes_table.item(item, tags=("black"))
    else:
        for item in routes_table.get_children():
            routes_table.item(item, tags=("black"))

    start_spinner(190, 133)
    
    with lock:
        active_threads = len(data)

    start_thread_for_route(data)


def start_thread_for_route(data):
    # Check if there are routes left to test
    if not data:
        return
    
    entry = data.pop(0)  # Get the next entry
    
    threading.Thread(target=test_route_and_update_ui, args=(entry,)).start()

    # Schedule the next thread execution
    routes_table.after(100, lambda: start_thread_for_route(data))


def test_route_and_update_ui(entry):
    global active_threads
    name, ip, ams_net_id, type_ = entry
    port = 851 if type_ == 'TC3' else 801

    connection_ok = test_connection(ams_net_id, port, name)
    routes_table.after(0, lambda: update_ui_with_result(name, connection_ok))

    # Decrement the thread counter and check if all threads are done
    with lock:
        active_threads -= 1
        if active_threads == 0:
            # Ensure that the spinner stops and the selection is removed after the UI update
            routes_table.after(0, lambda: stop_spinner())
            routes_table.after(0, lambda: routes_table.selection_remove(routes_table.selection()))
            routes_table.after(0, lambda: create_routes_button.config(state="normal"))


def test_connection(ams_net_id, port, name):
    plc = pyads.Connection(ams_net_id, port)
    try:
        # Open the connection
        plc.open()
        state = plc.read_state()
        print(f"PLC Status: {state}")

        # Check if the PLC is in RUN state (state[0] == 5)
        if state[0] == 5:
            print(f"Connection to {name} established successfully.")
            return True
        else:
            print(f"{name} is not in RUN state.")
            return False

    except pyads.ADSError as ads_error:
        print(f"ADS Error: {ads_error}")
        # messagebox.showerror('Error', f"ADS Error: {ads_error} Unable to connect to {name}")
        return False
    except Exception as e:
        print(f"Unexpected Exception: {e}")
        # messagebox.showerror('Error', f"Unexpected Exception: {e} Unable to connect to {name}")
        return False
    finally:
        # Ensure the connection is closed if it was successfully opened
        if plc.is_open:
            plc.close()
        time.sleep(0.5)


def update_ui_with_result(name, connection_ok):
    # Make sure to update the UI from the main thread
    def update():
        for item in routes_table.get_children():
            if routes_table.item(item, 'values')[0] == name:  # Assuming 'name' is in the first column
                color = 'green' if connection_ok else 'red'
                routes_table.item(item, tags=(color,))
                break

    # Use the `after` method to safely update the UI from the main thread
    routes_table.after(0, update)

def check_inputs():
    username = username_entry.get()
    password = password_entry.get()
    if not username:
        messagebox.showerror("Attention", "Add username!")
        return None
    if not password:
        messagebox.showerror("Attention", "Add password!")
        return None
    
    data = get_data_for_routes()
    if not data:
        messagebox.showerror("Attention", "Routes table is empty!")
        return None
    return data, username, password

def get_data_for_routes():
    data = []
    selected_items = routes_table.selection()
    if selected_items: 
        for item in selected_items:
            data.append(routes_table.item(item)["values"])
    else:
        data = get_table_data()
    print(data)
    return data
    
def get_table_data():
    rows = []
    for item in routes_table.get_children():
        rows.append(routes_table.item(item)["values"])
    return rows

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


def is_descendant(widget, parent):
    while widget:
        if widget == parent:
            return True
        widget = widget.master
    return False

def on_click(event):
    # print(f"x: {root.winfo_pointerx()}, y: {root.winfo_pointery()}")
    widget = event.widget
    if widget not in exceptions and not any(is_descendant(widget, exception) for exception in exceptions):
        routes_table.selection_remove(routes_table.selection())


############################# Spinner ###########################################3
def create_spinner_widget():
    global spinner_frame, spinner_canvas, spinner_arc

    # Create a frame to hold the spinner (fixed position in the layout)
    spinner_frame = tk.Frame(root, width=25, height=25, bg=root['bg'])  # Match frame bg to window bg

    # Create a canvas for the spinner with the same background color as the root window
    spinner_canvas = tk.Canvas(spinner_frame, width=25, height=25, bg=root['bg'], highlightthickness=0)
    spinner_canvas.pack()

    # Draw a rotating arc (spinner)
    spinner_arc = spinner_canvas.create_arc((2, 2, 22, 22), start=0, extent=90, width=4, outline='blue', style=tk.ARC)

    # Initially hide the spinner frame
    spinner_frame.place_forget()

def start_spinner(x, y):
    global running
    running = True  # Set the spinner running flag

     # Make the spinner visible
    spinner_frame.place(x=x, y=y)  # Adjust position as needed

    rotate_spinner()  # Start rotating the spinner

def stop_spinner():
    global running
    running = False  # Stop the spinner from running

    # Hide the spinner frame
    spinner_frame.place_forget()

def rotate_spinner():
    global spinner_arc
    if running:
        current_angle = spinner_canvas.itemcget(spinner_arc, 'start')
        new_angle = (float(current_angle) + 20) % 360  # Adjust rotation speed here
        spinner_canvas.itemconfig(spinner_arc, start=new_angle)
        spinner_canvas.after(50, rotate_spinner)  # Adjust the delay for rotation speed

############################# Set GUI icon ##########################
def set_icon():
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)
    else:
        print("Icon file not found.")

################################################################# Set up the GUI ######################################################################
root = tk.Tk()
root.title(f"Super Routes Creator {__version__}")

spinner_window = None
# Check if running as a script or frozen executable
if getattr(sys, 'frozen', False):
    icon_path = os.path.join(sys._MEIPASS, "./route.ico")
else:
    icon_path = os.path.abspath("./route.ico")
# root.iconbitmap(icon_path)

window_width = 430
window_lenght = 550
root.geometry(f"{window_width}x{window_lenght}")
root.minsize(window_width, window_lenght)
# root.resizable(True, True)

# Apply the icon after the window is initialized
root.after(100, set_icon)

# italic_font = font.Font(family="Segoe UI", size=10, slant="italic")

# Create a custom style for the LabelFrame with an italic font
style = ttk.Style()
style.configure("Custom.TLabelframe.Label", font=("Segoe UI", 10, "italic"))

# Create a style for the Entry widget
style.configure('Project.TEntry', foreground='black')
style.configure('Range.TEntry', foreground='black')
style.configure('BaseIP.TEntry', foreground='black')

style.configure('Placeholder.TEntry', foreground='grey')

# Customize the button's style for when it gains focus
style.configure("TButton", focuscolor="white", focusthickness=1)

frame_tc = tk.Frame(root)
frame_tc.grid(row=0, column=1, padx=5, pady=5)
optionTC = tk.StringVar(value="TC3")
tc2_radio = ttk.Radiobutton(frame_tc, text="TC2", variable=optionTC, value="TC2")
tc2_radio.grid(row=0, column=0, padx=0, pady=0, sticky='w')
tc3_radio = ttk.Radiobutton(frame_tc, text="TC3", variable=optionTC, value="TC3")
tc3_radio.grid(row=0, column=1, padx=0, pady=0, sticky='w')

frame_project = tk.Frame(root)
frame_project.grid(row=0, column=0, padx=5, pady=5, sticky='e')
label_project = ttk.Label(frame_project, text="Project number CC:")
label_project.grid(row=0, column=0, padx=0, pady=5)

entry_project = ttk.Entry(frame_project, style="Project.TEntry")#, fg="grey"
entry_project.grid(row=0, column=1, padx=5, pady=5)
create_placeholder(entry_project, "e.g., 1584", "Project.TEntry", "Placeholder.TEntry")
entry_project.bind("<KeyRelease>", validate_entry(entry_project, 'Project.TEntry', validate_project))

frame_range = tk.Frame(root)
frame_range.grid(row=1, column=0, padx=5, pady=5, sticky='e')
frame_lgv = tk.Frame(frame_range)
frame_lgv.grid(row=0, column=0, padx=1, pady=1)
optionLGV = tk.StringVar(value="LGV")
cb_radio = ttk.Radiobutton(frame_lgv, text="CB", variable=optionLGV, value="CB")
cb_radio.grid(row=0, column=0, padx=1, pady=1, sticky='w')
lgv_radio = ttk.Radiobutton(frame_lgv, text="LGV: ", variable=optionLGV, value="LGV")
lgv_radio.grid(row=0, column=1, padx=1, pady=1, sticky='w')

entry_lgv_range = ttk.Entry(frame_range, style="Range.TEntry")#, fg="grey")
entry_lgv_range.grid(row=0, column=1, padx=5, pady=5)
create_placeholder(entry_lgv_range, "e.g., 1-5,11-17,20-25", "Range.TEntry", "Placeholder.TEntry")
entry_lgv_range.bind("<KeyRelease>", validate_entry(entry_lgv_range, 'Range.TEntry', validate_range))

# Add a button to trigger table population
button_populate_table = ttk.Button(root, text=" Update Table ", 
                                #   bg="ghost white", 
                                  command=populate_table_from_inputs)
button_populate_table.grid(row=1, column=1, pady=10)
# button_design(button_populate_table)

frame_ip = tk.Frame(root)
frame_ip.grid(row=3, column=0, padx=5, pady=5, sticky='e')

# label_ip_help = tk.Label(frame_ip, text=" ? ", bd=2, relief='raised')
# label_ip_help = ttk.Label(frame_ip, text=" ? ")
style.configure("Help.TLabel", padding=2, relief="raised", font=("Segoe UI", 9))

label_ip_help = ttk.Label(frame_ip, text=" ? ", style="Help.TLabel")
label_ip_help.grid(row=0, column=0, padx=0, pady=5)
ToolTip(label_ip_help, "The IP of the first element of the range")


label_ip = ttk.Label(frame_ip, text="First IP: ")
label_ip.grid(row=0, column=1, padx=5, pady=5)

entry_base_ip = ttk.Entry(frame_ip, style="BaseIP.TEntry")#, fg="grey")
create_placeholder(entry_base_ip, "e.g., 172.20.3.11", "BaseIP.TEntry", "Placeholder.TEntry")
entry_base_ip.grid(row=0, column=2, padx=5, pady=5)
entry_base_ip.bind("<KeyRelease>", validate_entry(entry_base_ip, 'BaseIP.TEntry', validate_base_ip))


delete_table_button = ttk.Button(root, text=" Delete Table ", 
                                # bg="ghost white", 
                                command=delete_whole_table)
delete_table_button.grid(row=3, column=1, pady=10)
# button_design(delete_table_button)


frame_load = ttk.Labelframe(root, text="Load", labelanchor='nw', style="Custom.TLabelframe")
# frame_load = tk.LabelFrame(root, text="Load", labelanchor='nw', font=italic_font)
frame_load.grid(row=4, column=0, columnspan=1, padx=15, pady=5, sticky='w')
# Add a button to trigger the XML file selection and table population
load_xml_button = ttk.Button(frame_load, text=" StaticRoutes.xml ", 
                            # bg="ghost white", 
                            command=populate_table_from_xml)
load_xml_button.grid(row=1, column=0, padx=10, pady=5, sticky='ew')
# button_design(load_xml_button)

load_db3_button = ttk.Button(frame_load, text="     Config.db3     ", 
                            # bg="ghost white", 
                            command=populate_table_from_db3)
load_db3_button.grid(row=2, column=0, padx=10, pady=5, sticky='ew')
# button_design(load_db3_button)


# frame_login = ttk.Frame(root, bd=1, relief="groove")
frame_login = ttk.Labelframe(root, text="Router", labelanchor='nw', style="Custom.TLabelframe")
frame_login.grid(row=4, column=0, columnspan=2, padx=0, pady=5, sticky='e')

frame_user = tk.Frame(frame_login)
frame_user.grid(row=0, column=0, padx=5, pady=0)

username_label = ttk.Label(frame_user, text="Username:")
username_label.grid(row=0, column=0, padx=0, pady=5, sticky='e')

username_entry = ttk.Entry(frame_user, width=15)
username_entry.insert(0, "Administrator")
username_entry.grid(row=0, column=1, padx=5, pady=5)

frame_password = tk.Frame(frame_login)
frame_password.grid(row=1, column=0, padx=5, pady=0)

password_label = ttk.Label(frame_password, text="Password:")
password_label.grid(row=0, column=0, padx=0, pady=5, sticky='e')

password_entry = ttk.Entry(frame_password, show="*", width=15)
password_entry.grid(row=0, column=1, padx=5, pady=5)

test_routes_button = ttk.Button(frame_login, text="  Test Routes  ",
                                # bg="ghost white",
                                command=test_tc_routes)
test_routes_button.grid(row=0, column=1, padx=5, pady=5)
# button_design(test_routes_button)

create_routes_button = ttk.Button(frame_login, text="Create Routes",
                                # bg="ghost white",
                                command=create_tc_routes)
create_routes_button.grid(row=1, column=1, padx=5, pady=5)
# Disable it until test_tc_routes is done
create_routes_button.config(state="disabled")
# button_design(create_routes_button)



# Add a frame to hold the Treeview and the scrollbar
frame_table = tk.Frame(root)
frame_table.grid(row=5, columnspan=3, padx=15, pady=10)

# Add a Treeview to display the data
routes_table = ttk.Treeview(frame_table, columns=("Name", "Address", "NetId", "Type"), show="headings", height=10)
routes_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Define tags in your Treeview setup
routes_table.tag_configure('green', foreground='green')
routes_table.tag_configure('red', foreground='red')
routes_table.tag_configure('black', foreground='black')

setup_routes_table()

# Add a vertical scrollbar to the Treeview
vsb = ttk.Scrollbar(frame_table, orient="vertical", command=routes_table.yview)
vsb.pack(side=tk.RIGHT, fill=tk.Y)

# Configure the Treeview to use the scrollbar
routes_table.configure(yscrollcommand=vsb.set)

# Define the column widths
routes_table.column("Name", width=110, anchor='w')
routes_table.column("Address", width=110, anchor='w')
routes_table.column("NetId", width=120, anchor='w')
routes_table.column("Type", width=50, anchor='w')

routes_table.bind('<Delete>', delete_selected_record)
routes_table.bind('<Double-1>', on_double_click)
routes_table.bind('<<TreeviewSelect>>',  update_ssh_state)


frame_save_file = ttk.Labelframe(root, text="Save", labelanchor='nw', style="Custom.TLabelframe")
frame_save_file.grid(row=6, column=0, columnspan=4, padx=15, pady=5, sticky='w')
# save_label = tk.Label(frame_save_file, text="Save", font=italic_font)
# save_label.grid(row=0, column=0, padx=5, pady=0, sticky='w')
# Button to save the StaticRoutes.xml file
save_xml_button = ttk.Button(frame_save_file, text="   StaticRoutes   ", 
                        style="TButton", 
                        command=save_routes)
save_xml_button.grid(row=1, column=0, padx=5, pady=5)
# button_design(save_xml_button)

# Button to save the ControlCenter.xml file
save_cc_button = ttk.Button(frame_save_file, text="  ControlCenter   ", 
                            style="TButton", 
                            command=save_cc_xml)
save_cc_button.grid(row=1, column=1, padx=5, pady=5)
# button_design(save_cc_button)

# Button to save WinSCP.ini file
save_winscp_button = ttk.Button(frame_save_file, text="  WinSCP.ini   ", 
                            style="TButton", 
                            command=save_winscp_ini)
save_winscp_button.grid(row=1, column=2, padx=5, pady=5)
# button_design(save_winscp_button)

setup_tunnel_button = ttk.Button(frame_save_file, text="  Setup SSH   ", 
                            style="TButton", 
                            command=open_ssh_config_window_cond)
setup_tunnel_button.grid(row=1, column=3, padx=5, pady=5)


# Create the context menu
context_menu = tk.Menu(routes_table, tearoff=0)
# context_menu.add_command(label="Delete", command=delete_selected_record_from_menu)
context_menu.add_command(label="SSH Tunnel", command=create_ssh_tunnel)
context_menu.add_command(label="Open RDP", command=open_rdp_connection)

# Bind right-click to show the context menu
routes_table.bind("<Button-3>", show_context_menu)

exceptions = [routes_table, vsb, frame_login]
root.bind("<Button-1>", on_click)


# Create the spinner as part of the layout
create_spinner_widget()

check_twinCAT_version()

# Pupulate table the first time with current StaticRoutes.xml file
populate_table_from_xml("C:\\TwinCAT\\3.1\\Target\\StaticRoutes.xml")

root.mainloop()

# leer config.db3 y llenar tabla con eso - DONE

# agregar rutas de ads

# add local route by default

# DELETE
# [Configuration\LastFingerprints]
# 172.20.2.68=20022:ssh=ecdsa-sha2-nistp384%20384%20iXnY+SMyoQRSUxJMzgWWA+yadddMZqqgM4dLPp/uHhs
# 172.20.2.68:20022:ssh=ecdsa-sha2-nistp384%20384%20iXnY+SMyoQRSUxJMzgWWA+yadddMZqqgM4dLPp/uHhs

# Tener la posibilidad de borrar más rows al seleccionar shift o control

# Routes can be created even if static routes file is not updated (TwinCAT is not restarted yet). But after routes are created TwinCAT should be restarted to have the comm



# Add PuTTY sessions, first check if it is installed, if not, popup to show is not installed, if yes, create all the sessions on the registry