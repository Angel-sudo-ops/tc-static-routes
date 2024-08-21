import ctypes

# Load the DLL
ads_dll = ctypes.CDLL('./CRADSDriver.dll')

# Define the TwinCATCom class
class TwinCATCom(ctypes.Structure):
    _fields_ = [
        ("DisableSubScriptions", ctypes.c_bool),
        ("Password", ctypes.c_char_p),
        ("PollRateOverride", ctypes.c_int),
        ("TargetAMSNetID", ctypes.c_char_p),
        ("TargetAMSPort", ctypes.c_ushort),
        ("TargetIPAddress", ctypes.c_char_p),
        ("UserName", ctypes.c_char_p),
        ("UseStaticRoute", ctypes.c_bool),
    ]

# Instantiate the structure (this initializes it in Python)
twincatCom = TwinCATCom()

# Example of setting fields
twincatCom.DisableSubScriptions = True
twincatCom.Password = b"your_password"
twincatCom.TargetAMSNetID = b"10.209.80.202.1.1"
twincatCom.TargetIPAddress = b"10.209.80.202"
twincatCom.TargetAMSPort = 851  # Replace with the actual port number you need
twincatCom.UserName = b"Administrator"
twincatCom.UseStaticRoute = True

# Define the CreateRoute method
ads_dll.CreateRoute.argtypes = [ctypes.POINTER(TwinCATCom), ctypes.c_char_p, ctypes.c_char_p]
ads_dll.CreateRoute.restype = ctypes.c_char_p  # Adjust based on the actual return type

# Call the method
try:
    result = ads_dll.CreateRoute(ctypes.byref(twincatCom), b"Name", b"10.209.80.202")
    print(f"Route created successfully, result: {result.decode('utf-8')}")
except Exception as e:
    print(f"Failed to create route: {e}")
