import pyads

def get_local_ams_netid():
    return pyads.get_local_address()

# Example usage
local_ams_netid = get_local_ams_netid()
print(f"Local AMS NetID: {local_ams_netid}")