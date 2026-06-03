import ctypes 

dll = ctypes.CDLL(r"D:\Documents\PythonPrograms\StaticRoutesCreator\adslib.dll")
port = dll.AdsPortOpenEx()
print(f"port: {port}")
dll.AdsPortCloseEx(port)