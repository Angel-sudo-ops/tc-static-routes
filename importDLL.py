import clr
import System
from System import Activator
from System import Type
from System.Reflection import BindingFlags
from System.Net import IPAddress
from System import Array

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

    # Set m_TargetAMSNetID to a valid value before proceeding
    m_TargetAMSNetID_field = twin_cat_type.GetField("m_TargetAMSNetID", BindingFlags.NonPublic | BindingFlags.Instance)
    if m_TargetAMSNetID_field:
        m_TargetAMSNetID_field.SetValue(twincat_com, "10.209.80.202.1.1")
        print("m_TargetAMSNetID set successfully.")

    # Manually invoke CreateDLLInstance method to ensure initialization
    create_dll_instance_method = twin_cat_type.GetMethod(
        "CreateDLLInstance",
        BindingFlags.NonPublic | BindingFlags.Instance
    )
    if create_dll_instance_method:
        print("Calling CreateDLLInstance directly...")
        create_dll_instance_method.Invoke(twincat_com, None)
        print("CreateDLLInstance called successfully.")
    else:
        print("Failed to locate CreateDLLInstance method")

    # Check MyDLLInstance value after attempting to invoke CreateDLLInstance
    my_dll_instance_field = twin_cat_type.GetField("MyDLLInstance", BindingFlags.NonPublic | BindingFlags.Instance)
    if my_dll_instance_field:
        my_dll_instance_value = my_dll_instance_field.GetValue(twincat_com)
        print(f"MyDLLInstance value after CreateDLLInstance: {my_dll_instance_value}")
    
    # Check the DLL dictionary and manually add an entry if necessary
    dll_field = twin_cat_type.GetField("DLL", BindingFlags.NonPublic | BindingFlags.Static)
    if dll_field:
        dll_dict = dll_field.GetValue(None)
        print(f"DLL Dictionary contains {len(dll_dict)} items after CreateDLLInstance")

        # If the DLL dictionary doesn't contain the key, manually add it
        if my_dll_instance_value not in dll_dict:
            ads_for_twincat_ex_type = Type.GetType("TwinCATAds.ADSforTwinCATEx, " + assembly_name)
            if ads_for_twincat_ex_type:
                ads_for_twincat_ex = Activator.CreateInstance(ads_for_twincat_ex_type)
                dll_dict[my_dll_instance_value] = ads_for_twincat_ex
                print(f"Manually added DLL entry for key {my_dll_instance_value}")
            else:
                print("Failed to create ADSforTwinCATEx instance.")

    # Check the DLL dictionary after manual insertion
    if dll_field:
        dll_dict = dll_field.GetValue(None)
        print(f"DLL Dictionary contains {len(dll_dict)} items after manual insertion")

    # Now proceed with setting properties and calling CreateRoute
    twincat_com.DisableSubScriptions = True
    twincat_com.Password = "1"
    twincat_com.PollRateOverride = 500
    twincat_com.TargetAMSNetID = "10.209.80.202.1.1"  # Ensure this matches the AMS Net ID set earlier
    twincat_com.TargetIPAddress = "10.209.80.202"
    twincat_com.TargetAMSPort = 851
    twincat_com.UserName = "Administrator"
    twincat_com.UseStaticRoute = True

    local_ip = IPAddress.Parse("10.209.80.202")

    # Debugging before calling CreateRoute
    print("Calling CreateRoute...")
    try:
        result = twincat_com.CreateRoute("Name", local_ip)
        print(f"Route created successfully, result: {result}")
    except Exception as e:
        print(f"Error during CreateRoute invocation: {e}")
