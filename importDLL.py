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

    # Set m_TargetAMSNetID to a valid value before invoking Write
    m_TargetAMSNetID_field = twin_cat_type.GetField("m_TargetAMSNetID", BindingFlags.NonPublic | BindingFlags.Instance)
    if m_TargetAMSNetID_field:
        m_TargetAMSNetID_field.SetValue(twincat_com, "10.209.80.202.1.1")
        print("m_TargetAMSNetID set successfully.")

    # Check MyDLLInstance value before proceeding
    my_dll_instance_field = twin_cat_type.GetField("MyDLLInstance", BindingFlags.NonPublic | BindingFlags.Instance)
    if my_dll_instance_field:
        my_dll_instance_value = my_dll_instance_field.GetValue(twincat_com)
        print(f"Initial MyDLLInstance value: {my_dll_instance_value}")
    
    # Check the DLL dictionary
    dll_field = twin_cat_type.GetField("DLL", BindingFlags.NonPublic | BindingFlags.Static)
    if dll_field:
        dll_dict = dll_field.GetValue(None)
        print(f"DLL Dictionary contains {len(dll_dict)} items before Write")

    # Attempt to invoke Write method
    string_type = clr.GetClrType(str)
    array_string_type = clr.GetClrType(System.Array[System.String])
    ushort_type = clr.GetClrType(System.UInt16)

    write_method = twin_cat_type.GetMethod(
        "Write",
        BindingFlags.Public | BindingFlags.Instance,
        None,
        Array[Type]([string_type, array_string_type, ushort_type]),
        None
    )

    if write_method:
        print("Calling Write method to trigger CreateDLLInstance...")
        start_address = "SomeAddress"  # Example, replace with the actual address you need
        data_to_write = Array[str](["data1", "data2"])  # Example data, adjust accordingly
        number_of_elements = System.UInt16(2)  # Example, adjust as needed

        try:
            write_method.Invoke(twincat_com, [start_address, data_to_write, number_of_elements])
            print("Write method called successfully.")
        except Exception as e:
            print(f"Error during Write method invocation: {e}")

    # Check the DLL dictionary after Write
    if dll_field:
        dll_dict = dll_field.GetValue(None)
        print(f"DLL Dictionary contains {len(dll_dict)} items after Write")

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
    result = twincat_com.CreateRoute("Name", local_ip)
    print(f"Route created successfully, result: {result}")
