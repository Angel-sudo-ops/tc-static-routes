import clr
from System.Net import IPAddress

# Assuming you've already added the reference to the CRADSDriver.dll
clr.AddReference('CRADSDriver')

from TwinCATAds import TwinCATCom

# Instantiate TwinCATCom
twincat_com = TwinCATCom()

twincat_com.CreateDLLInstance()

# Set other properties as before
twincat_com.DisableSubScriptions = True
twincat_com.Password = "1"
twincat_com.TargetAMSNetID = "10.209.80.202.1.1"
twincat_com.TargetIPAddress = "10.209.80.202"
twincat_com.TargetAMSPort = 851
twincat_com.UserName = "Administrator"
twincat_com.UseStaticRoute = True

# Create the IPAddress object
local_ip = IPAddress.Parse("10.209.80.202")  # Convert the IP string to an IPAddress object

# Call CreateRoute with the correct types
result = twincat_com.CreateRoute("Name", local_ip)

# Output the result
print(f"Route created successfully, result: {result}")
