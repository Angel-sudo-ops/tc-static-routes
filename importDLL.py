import clr
import System
from System import Activator
from System import Type
from System.Reflection import BindingFlags
from System.Net import IPAddress

# Load the assembly explicitly if not already done
clr.AddReference('CRADSDriver')

# Use the fully qualified name, including the assembly name, if necessary
assembly_name = "CRADSDriver"
type_name = "TwinCATAds.TwinCATCom, " + assembly_name

# Get the type using the fully qualified name
twin_cat_type = Type.GetType(type_name)

if twin_cat_type is None:
    print(f"Failed to find type '{type_name}'")
else:
    # Instantiate the TwinCATCom object using Activator
    twincat_com = Activator.CreateInstance(twin_cat_type)

    # Get the EndInit method via reflection
    end_init_method = twin_cat_type.GetMethod(
        "System.ComponentModel.ISupportInitialize.EndInit",
        BindingFlags.NonPublic | BindingFlags.Instance
    )

    # Invoke the EndInit method if found
    if end_init_method:
        print("Calling EndInit...")
        end_init_method.Invoke(twincat_com, None)
        print("EndInit called successfully.")
    else:
        print("Failed to locate EndInit method")

    # Consider other initialization steps (hypothetical)
    # If there's another method that should be called, call it here
    # example: twincat_com.OtherInitializationMethod()

    # Now proceed with setting properties and calling CreateRoute
    twincat_com.DisableSubScriptions = True
    twincat_com.Password = "your_password"
    twincat_com.TargetAMSNetID = "10.209.80.202.1.1"
    twincat_com.TargetIPAddress = "10.209.80.202"
    twincat_com.TargetAMSPort = 851
    twincat_com.UserName = "Administrator"
    twincat_com.UseStaticRoute = True

    local_ip = IPAddress.Parse("10.209.80.202")

    # Debugging before calling CreateRoute
    print("Calling CreateRoute...")
    result = twincat_com.CreateRoute("Name", local_ip)
    print(f"Route created successfully, result: {result}")
