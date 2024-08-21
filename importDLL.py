import clr
from System.Net import IPAddress
from System.ComponentModel import ISupportInitialize

# Load the necessary assembly
clr.AddReference('CRADSDriver')  # Adjust based on your DLL

from TwinCATAds import TwinCATCom

# Instantiate the TwinCATCom object
twincat_com = TwinCATCom()

# Directly call the EndInit method if accessible
try:
    twincat_com.EndInit()  # Call EndInit directly
except AttributeError:
    # If EndInit isn't directly accessible, we need to use reflection or another approach
    print("EndInit method not directly accessible")

# Proceed with setting properties as before
twincat_com.DisableSubScriptions = True
twincat_com.Password = "your_password"
twincat_com.TargetAMSNetID = "10.209.80.202.1.1"
twincat_com.TargetIPAddress = "10.209.80.202"
twincat_com.TargetAMSPort = 851
twincat_com.UserName = "Administrator"
twincat_com.UseStaticRoute = True

# Create the IPAddress object for the second argument
local_ip = IPAddress.Parse("10.209.80.202")

# Call CreateRoute with the correct types
result = twincat_com.CreateRoute("Name", local_ip)

# Output the result
print(f"Route created successfully, result: {result}")
