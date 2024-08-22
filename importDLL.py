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

    # Set a valid AMS Net ID before invoking CreateDLLInstance
    m_TargetAMSNetID_field = twin_cat_type.GetField("m_TargetAMSNetID", BindingFlags.NonPublic | BindingFlags.Instance)
    if m_TargetAMSNetID_field:
        m_TargetAMSNetID_field.SetValue(twincat_com, "10.209.80.202.1.1")  # Example AMS Net ID, adjust as needed
        print("m_TargetAMSNetID set successfully.")
    else:
        print("Failed to retrieve m_TargetAMSNetID field")

    # Get the private CreateDLLInstance method via reflection
    create_dll_instance_method = twin_cat_type.GetMethod(
        "CreateDLLInstance",
        BindingFlags.NonPublic | BindingFlags.Instance
    )

    # Invoke the CreateDLLInstance method if found
    if create_dll_instance_method:
        print("Calling CreateDLLInstance directly...")
        create_dll_instance_method.Invoke(twincat_com, None)
        print("CreateDLLInstance called successfully.")
    else:
        print("Failed to locate CreateDLLInstance method")

    # Additional debugging to see if MyDLLInstance has been set
    my_dll_instance_field = twin_cat_type.GetField("MyDLLInstance", BindingFlags.NonPublic | BindingFlags.Instance)
    if my_dll_instance_field:
        my_dll_instance_value = my_dll_instance_field.GetValue(twincat_com)
        print(f"MyDLLInstance value: {my_dll_instance_value}")
    else:
        print("Failed to retrieve MyDLLInstance field")

    # Now proceed with setting properties and calling CreateRoute
    twincat_com.DisableSubScriptions = True
    twincat_com.Password = "your_password"
    twincat_com.TargetAMSNetID = "10.209.80.202.1.1"  # Ensure this matches the AMS Net ID set earlier
    twincat_com.TargetIPAddress = "10.209.80.202"
    twincat_com.TargetAMSPort = 851
    twincat_com.UserName = "Administrator"
    twincat_com.UseStaticRoute = True

    local_ip = IPAddress.Parse("10.209.80.202")

    # Debugging before calling CreateRoute
    print("Calling CreateRoute...")
    result = twincat_com.CreateRoute("Name", local_ip)
    print(f"Route created successfully, result: {result}")
