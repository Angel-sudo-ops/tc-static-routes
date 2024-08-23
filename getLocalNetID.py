import clr
import System
from System import Activator
from System import Type
from System.Reflection import BindingFlags
from System.Net import IPAddress
import os
import sys

def get_local_ams_netid():
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

    # Use the fully qualified name, including the assembly name, if necessary
    assembly_name = "CRADSDriver"
    type_name = "TwinCATAds.ADSforTwinCAT, " + assembly_name

    # Get the type using the fully qualified name
    ads_twincat_type = Type.GetType(type_name)

    if ads_twincat_type is None:
        print(f"Failed to find type '{type_name}'")
        return None

    # Instantiate the TwinCATCom object using Activator
    ads_twincat_com = Activator.CreateInstance(ads_twincat_type)

    return ads_twincat_com.get_MyAMSNetID()





# Example usage
local_ams_netid = get_local_ams_netid()
print(f"Local AMS NetID: {local_ams_netid}")