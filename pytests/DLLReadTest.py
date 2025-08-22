import clr
import sys
from pathlib import Path
import os
import System

def load_dlls():
    if getattr(sys, 'frozen', False):
        # If the application is frozen (bundled by PyInstaller), get the path of the executable
        base_path = sys._MEIPASS
    else:
        # If running as a script, use the current directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    # Define the path to the DLL folder relative to the base path
    dll_folder = os.path.join(base_path, "dll")

    # Check if the folder exists
    if os.path.exists(dll_folder) and os.path.isdir(dll_folder):
        print(f"Found DLL folder: {os.path.abspath(dll_folder)}")
    else:
        print(f"DLL folder not found: {os.path.abspath(dll_folder)}")

    # Load the DLLs from the folder
    for dll in os.listdir(dll_folder):
        if dll.endswith(".dll"):
            dll_path = os.path.join(dll_folder, dll)
            try:
                clr.AddReference(dll_path)
                print(f"Loaded {dll}")
            except Exception as e:
                print(f"Failed to load {dll}: {e}")

load_dlls()

# Try to list all types in the assembly and catch detailed loader exceptions
try:
    for assembly in System.AppDomain.CurrentDomain.GetAssemblies():
        # print(assembly.FullName)
        try:
            for typ in assembly.GetTypes():
                print(f"Type: {typ.FullName}")
        except System.Reflection.ReflectionTypeLoadException as ex:
            # Print detailed loader exceptions
            print("LoaderExceptions:")
            for loader_exception in ex.LoaderExceptions:
                print(loader_exception)
                print("\n")
except Exception as e:
    print(f"General Error: {e}")
    

# Import the necessary classes from the TwinCAT.Ads namespace
from TwinCAT.Ads import AdsClient, AmsAddress, AmsNetId
port = 851
known_ams_net_id = "172.20.1.67.1.1"  # Example AMS Net ID
target_address = AmsAddress(known_ams_net_id, port)

# Create an instance of AdsClient
client = AdsClient()

try:
    print(f"AMS Net Id : {AmsNetId.Local}")
    print(f"Local Host: {AmsNetId.LocalHost}")
    print(f"Local: {AmsNetId.IsLocal}")
    print(f"Loop: {AmsNetId.IsLoopback}")
    print(f"Loop: {AmsNetId.Empty}")
except Exception as e:
    print(f"Error: {e}")

try:
   
    # Get the local AMS Net ID
    local_address = client.get_Address()
    local_net_id = local_address.NetId.ToString()
    # Print the AMS Net ID
    print(f"Local AMS Net ID: {local_net_id}")

except Exception as e:
    print(f"Error retrieving AMS Net ID: {e}")

try:
    # Connect to the target AMS address
    # client.Connect(target_address)
    # client.Connect(target_address)
    client.Connect(known_ams_net_id, port)

    local_address = client.get_Address()
    local_net_id = local_address.NetId.ToString()
    print(f"Local AMS Net ID: {local_net_id}")
    # local_name = client.get_Name()
    # print(f"Name: {local_name}")

    if port == 851:
        test = client.CreateVariableHandle("MAIN.startupOk")
    else:
        test = client.CreateVariableHandle(".StartIp_ok")

    resp = client.ReadAny(test, type(bool))
    # Further operations can be attempted here
    device_info = client.ReadDeviceInfo()
    print(f"Device Info: {device_info}")
    
except Exception as e:
    print(f"Error: {e}")
finally:
    # Disconnect from the TwinCAT system
    client.Disconnect()


import platform
from TwinCATAds import TwinCATCom, ADSforTwinCATEx
from System.Net import IPAddress


ads_for_tc = ADSforTwinCATEx()

netid = str(AmsNetId.Local).split('.')
netid_ip = '.'.join(netid[:4])

local_ip = IPAddress.Parse(netid_ip)

twincat_comm = TwinCATCom()

# twincat_comm.DLL

# twincat_comm.CreateDLLInstance()
# dll_field = twincat_comm.DLL
# dll_field[twincat_comm.MyDLLInstance] = ads_for_tc

twincat_comm.DisableSubScriptions = True
twincat_comm.Password = "1"
twincat_comm.PollRateOverride = 500
twincat_comm.TargetAMSNetID = "172.20.2.81.1.1"
twincat_comm.TargetAMSPort = 801 #851
twincat_comm.TargetIPAddress = "172.20.2.81.1.1"
twincat_comm.UserName = "Administrator"
twincat_comm.UseStaticRoute = True

print(twincat_comm.TargetIPAddress)

print(local_ip)
print(platform.node())
twincat_comm.CreateRoute(str(platform.node()), local_ip)

