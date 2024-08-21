import ctypes

# dll_path = './CRADSDriver.dll'

# def readDLL(path):
#     try:
#         adsDriver = ctypes.CDLL(path)
#         print("DLL loaded successfully!")
#     except OSError as e:
#         print(f"Failed to load DLL: {e}")
    
#     adsDriver.__getattr__

# readDLL(dll_path)

ads_dll = ctypes.CDLL('./CRADSDriver.dll')

ads_dll.CreateRoute.argtypes = [ctypes.c_char_p, ctypes.c_char_p]
ads_dll.CreateRoute.restype = ctypes.c_char_p

# class TwincatCom(ctypes.Structure):
#     _fields_ = [
    
#     ]

# twincatComm = ads_dll

# # Set properties (assuming there are methods to set these or direct assignment is possible)
# twincatComm.DisableSubScriptions = True
# twincatComm.Password = ctypes.create_string_buffer(b"your_password")
# twincatComm.PollRateOverride = 500
# twincatComm.TargetAMSNetID = ctypes.create_string_buffer(b"10.209.80.202.1.1")
# twincatComm.TargetAMSPort = ctypes.c_uint16(801)
# twincatComm.TargetIPAddress = ctypes.create_string_buffer(b"10.209.80.202")
# twincatComm.UserName = ctypes.create_string_buffer(b"Administrator")
# twincatComm.UseStaticRoute = True


# try:
#     response = twincatComm.CreateRoute(
#         ctypes.create_string_buffer(b"Name"),
#         ctypes.create_string_buffer(b"10.209.80.202")
#     )
#     print(f"Connection successful, response: {response}")
# except Exception as e:
#     # Handle the PLCDriverException
#     print(f"Failed to connect: {str(e)}")