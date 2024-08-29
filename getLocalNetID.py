import clr
import System
from System import Activator
from System import Type
from System.Reflection import BindingFlags
from System.Net import IPAddress
import os
import sys
import platform
from System.Reflection import Assembly


# def load_dll():
#     # Determine if the application is running as a standalone executable
#     if getattr(sys, 'frozen', False):
#         # If the application is frozen (bundled by PyInstaller), get the path of the executable
#         base_path = sys._MEIPASS
#     else:
#         # If running as a script, use the current directory
#         base_path = os.path.dirname(os.path.abspath(__file__))

#     # Construct the full path to the DLL
#     dll_path = os.path.join(base_path, "CRADSDriver.dll")

#     # Load the assembly
#     clr.AddReference(dll_path)

# def get_local_ams_netid():
#     # Use the fully qualified name, including the assembly name, if necessary
#     assembly_name = "CRADSDriver"
#     type_name = "TwinCATAds.ADSforTwinCAT, " + assembly_name

#     # Get the type using the fully qualified name
#     ads_twincat_type = Type.GetType(type_name)

#     if ads_twincat_type is None:
#         print(f"Failed to find type '{type_name}'")
#         return None

#     # Instantiate the TwinCATCom object using Activator
#     ads_twincat = Activator.CreateInstance(ads_twincat_type)

#     return ads_twincat.get_MyAMSNetID()
    


def load_dll_v2():
    # Determine if the application is running as a standalone executable
    if getattr(sys, 'frozen', False):
        # If the application is frozen (bundled by PyInstaller), get the path of the executable
        base_path = sys._MEIPASS
    else:
        # If running as a script, use the current directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    # Construct the full path to the DLL
    dll_path = os.path.join(base_path, "TwinCAT.Ads.Abstractions.dll")

    # Load the assembly
    # clr.AddReference(dll_path)
    clr.AddReference(dll_path)

    # Use the fully qualified name, including the assembly name, if necessary
    assembly_name = 'TwinCAT.Ads.Abstractions'
    type_name = 'TwinCAT.Ads.AMSNetId, ' + assembly_name

    # Get the type using the fully qualified name
    # ads_type = Type.GetType(type_name)
    ads_type = Type.GetType("TwinCAT.Ads.AmsNetId, TwinCAT.Ads")

    if ads_type is None:
        print(f"Failed to find type '{type_name}'")
        return None

    # Instantiate the TwinCATCom object using Activator
    ads_twincat = Activator.CreateInstance(ads_type)

    # return ads_twincat.get_Address()
    return ads_type.AmsNetId.ToString()

# Example usage
# load_dll()
# local_ams_netid = get_local_ams_netid()
# print(f"Local AMS NetID: {local_ams_netid}")

# load_dll_v2()
# local_ams_netid_v2 = get_local_ams_netid_v2()
local_ams_netid_v2 = load_dll_v2()
print(f"Local AMS NetID: {local_ams_netid_v2}")

system_name = platform.system()
node_name = platform.node()
release = platform.release()
version = platform.version()
machine = platform.machine()
processor = platform.processor()

print(f"System Name: {system_name}")
print(f"Node Name: {node_name}")
print(f"Release: {release}")
print(f"Version: {version}")
print(f"Machine: {machine}")
print(f"Processor: {processor}")