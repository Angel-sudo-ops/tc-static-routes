import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import xml.etree.ElementTree as ET
from xml.dom import minidom
import sqlite3

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
    import re
    return [int(c) if c.isdigit() else c for c in re.split('(\d+)', text)]


#################################### Read config.db3 #######################################
def read_db3_file(db3_file_path, table_name):
    try:
        # Connect to the .db3 file
        conn = sqlite3.connect(db3_file_path)
        cursor = conn.cursor()

        # Check if the table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
        if not cursor.fetchone():
            messagebox.showerror("Error", f"Table '{table_name}' does not exist in the database.")
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
root.title("Static Routes XML Creator 1.4")

#Disable resizing
root.resizable(False, False)
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

# Add a frame to hold the Treeview and the scrollbar
frame_table = tk.Frame(root)
frame_table.grid(row=4, columnspan=3, padx=15, pady=10)

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


frame_xml = tk.Frame(root)
frame_xml.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
frame_load = tk.Frame(frame_xml)
frame_load.grid(row=0, column=0, padx=10, pady=10)
# Add a button to trigger the XML file selection and table population
button_load_xml = tk.Button(frame_load, text="Load StaticRoutes.xml", 
                            bg="ghost white", 
                            command=populate_table_from_xml)
button_load_xml.grid(row=0, column=0, padx=5, pady=5)
button_design(button_load_xml)

button_load_db3 = tk.Button(frame_load, text="     Load Config.db3     ", 
                            bg="ghost white", 
                            command=populate_table_from_db3)
button_load_db3.grid(row=1, column=0, padx=5, pady=5)
button_design(button_load_db3)

# Button to save the StaticRoutes.xml file
save_button = tk.Button(frame_xml, text="Save StaticRoutes.xml", 
                        bg="ghost white", 
                        command=save_routes_xml)
save_button.grid(row=0, column=1, padx=10, pady=10)
button_design(save_button)

# Button to save the ControlCenter.xml file
create_cc_button = tk.Button(frame_xml, text="Save ControlCenter file", 
                             bg="ghost white", 
                             command=save_cc_xml)
create_cc_button.grid(row=0, column=2, padx=10, pady=10)
button_design(create_cc_button)



# Create the context menu
context_menu = tk.Menu(treeview, tearoff=0)
context_menu.add_command(label="Delete", command=delete_selected_record_from_menu)

# Bind right-click to show the context menu
treeview.bind("<Button-3>", show_context_menu)

root.mainloop()

# leer config.db3 y llenar tabla con eso - DONE

# agregar rutas de ads

# add local route by default