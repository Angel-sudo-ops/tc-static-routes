import clr
import threading
import os
import sys
import platform

import System
from System import Activator 
from System import Type
from System.Reflection import BindingFlags
from System.Net import IPAddress


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
                dll_dict[my_dll_instance_value] = twincat_com  #ads_for_twincat_ex
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

# Example usage with your data
data = [['CC1965_LGV18', '172.20.2.68', '172.20.2.68.1.1', 'TC2'], 
        ['CC1965_LGV17', '172.20.2.67', '172.20.2.67.1.1', 'TC3'], 
       ]

# Call the function with the data, username, and password
create_routes_from_data(data, "Administrator", "1")
