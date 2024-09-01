import winreg

def get_ams_net_id():
    paths = [
        r"SOFTWARE\Beckhoff\TwinCAT3\System",  # TwinCAT 3
        r"SOFTWARE\Beckhoff\TwinCAT\System",    # TwinCAT 2
        r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT3\System",  # TwinCAT 3
        r"SOFTWARE\WOW6432Node\Beckhoff\TwinCAT\System"    # TwinCAT 2
    ]
    
    for path in paths:
        try:
            # Attempt to open the registry key
            reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
            ams_net_id, _ = winreg.QueryValueEx(reg_key, "AmsNetId")
            return ams_net_id
        except FileNotFoundError:
            continue  # If this path doesn't exist, try the next one
        except Exception as e:
            print(f"Failed to get AMS Net ID from registry path {path}: {e}")
            return None
    
    print("AMS Net ID not found in the registry.")
    return None

def format_ams_net_id(raw_net_id):
    # Convert the byte array to an IP-like string format
    return '.'.join(str(byte) for byte in raw_net_id)

# Get and print the AMS Net ID
raw_ams_net_id = get_ams_net_id()
if raw_ams_net_id:
    formatted_ams_net_id = format_ams_net_id(raw_ams_net_id)
    print(f"Local AMS Net ID: {formatted_ams_net_id}")
else:
    print("Failed to retrieve the AMS Net ID.")