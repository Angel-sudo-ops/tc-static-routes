import winreg
import struct

def save_route_tc2(route_name, net_id, address, flags=0, timeout=0, transport_type=1):
    try:
        reg_path = r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT\Remote"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_SET_VALUE)

        # Create a new subkey for the route using route_name (e.g., "CC1393_01")
        route_key = winreg.CreateKey(key, route_name)
        
        # Set the route details
        winreg.SetValueEx(route_key, "Address", 0, winreg.REG_SZ, address)

        # Convert the AMS Net ID from string (e.g., "5.72.127.151.1.1") to binary format
        net_id_bytes = struct.pack('6B', *[int(x) for x in net_id.split('.')])
        winreg.SetValueEx(route_key, "AmsNetId", 0, winreg.REG_BINARY, net_id_bytes)

        # Set other values (Flags, Timeout, TransportType)
        winreg.SetValueEx(route_key, "Flags", 0, winreg.REG_DWORD, flags)
        winreg.SetValueEx(route_key, "Timeout", 0, winreg.REG_DWORD, timeout)
        winreg.SetValueEx(route_key, "TransportType", 0, winreg.REG_DWORD, transport_type)
        
        winreg.CloseKey(route_key)
        winreg.CloseKey(key)
        print(f"Route {route_name} added to TwinCAT2 registry.")
    except Exception as e:
        print(f"Failed to save route: {e}")

# Example usage:
save_route_tc2("CC1200_LGV01", "10.12.31.101.1.1", "192.168.1.2")
save_route_tc2("CC1200_LGV02", "10.12.31.102.1.1", "192.168.1.2")