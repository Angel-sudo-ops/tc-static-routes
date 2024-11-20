import winreg
import struct

def check_twinCAT_version():
    try:
        # Check if TwinCAT3 is installed
        tc3_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT3")
        print("TwinCAT3 is installed.")
        winreg.CloseKey(tc3_key)
        return "TC3"
    except FileNotFoundError:
        pass
    
    try:
        # Check if TwinCAT2 is installed
        tc2_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT")
        print("TwinCAT2 is installed.")
        winreg.CloseKey(tc2_key)
        return "TC2"
    except FileNotFoundError:
        print("Neither TwinCAT3 nor TwinCAT2 are installed.")
        return None

def save_route_tc2(route_name, net_id, address, flags=0, timeout=0, transport_type=1):
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
    except PermissionError:
        print("Access denied: Please run the application as administrator.")
    except Exception as e:
        print(f"Failed to save route: {e}")

# Example usage:
# save_route_tc2("CC1200_LGV01", "10.40.10.71.1.1", "10.40.10.71")
# save_route_tc2("CC1200_LGV02", "10.40.10.72.1.1", "10.40.10.72")

check_twinCAT_version()