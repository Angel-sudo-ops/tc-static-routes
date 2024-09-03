import pyads
import platform

ams_port_TC3 = pyads.PORT_TC3PLC1    # Default port for TwinCAT 3 PLC
ams_port_TC2 = pyads.PORT_TC2PLC1

username = "Administrator"
password = "1"

system_name = platform.node()

def create_connection(host_name, plc_ip_address, username, password, port):
    ams_net_id = plc_ip_address + ".1.1"
    plc = pyads.Connection(ams_net_id, port)
    plc.open()
    if plc.is_open:
        print("PLC connection open")
        local_ams_addres = plc.get_local_address()
        local_ams_addres_net_id = local_ams_addres.netid
        print(local_ams_addres_net_id)
        plc.close()
    else:
        print("Unable to get AMS Net id")

    try:
        add_route = pyads.add_route_to_plc(
                local_ams_addres_net_id,
                host_name,
                plc_ip_address,
                username,
                password
            )
    except Exception as e:
        print(f"Exception: {e}")
        add_route = False
    if add_route:
        print("Route added sucessfully!!")
    else:
        print(":(")

def test_connection(ams_net_id, port):
    plc = pyads.Connection(ams_net_id, port)
    try:
        # Open the connection
        plc.open()
        # Check if the connection is established
        if plc.is_open:
            print("PLC connection open")
            ams_addres = plc.get_local_address()
            ams_addres_net_id = ams_addres.netid
            print(ams_addres_net_id)

            state = plc.read_state()
            print(f"PLC Status: {state}")
            if state[0] == 5: #PLC in run
                print("Connection to PLC established successfully.") 
                plc.close()                                
                return True
            else:
                print("Failed to establish connection to PLC.")
        else:
            print("Failed to establish connection to PLC.")
        # return False
    except pyads.ADSError as ads_error:
        print(f"ADS Error: {ads_error}")
        # return False
    
    except Exception as e:
        print(f"Exception: {e}")
        # return False
    plc.close()
    return False


def check_plc_connection(ams_net_id, port):
    try:
        # Open the ADS port
        pyads.open_port()

        # Create a connection object
        plc = pyads.Connection(ams_net_id, port)
        
        # Attempt to read the PLC state
        state = plc.read_state()
        print("PLC State:", state)
        return True
        
    except Exception as e:
        print("PLC connection failed:", e)
        return False
    
def get_local_ams_id():
    ams_net_id = None
    try:
        # Create a connection object
        plc = pyads.Connection('0.0.0.0.0.0', pyads.PORT_TC3PLC1)
        plc.open()

        ams_net_id = plc.get_local_address().netid
        
    # except pyads.ADSError as e:
    #     print(f"ADS Error: {e} \nCheck if TwiCAT on local machine is running")
    # except (OSError, IOError) as e:
    #     print(f"OS or I/O Error: {e} \nCheck if TwiCAT on local machine is running")
    except Exception as e:
        print(f"Unexpected error: {e} \nCheck if TwiCAT on local machine is running")
    finally:
        if plc.is_open:
            plc.close()

    return ams_net_id

# Define ADS parameters
plc_ip_address = "172.20.2.66"      # Replace with your PLC's IP address
plc_ams_net_id = plc_ip_address + ".1.1" # Replace with your PLC's AMS Net ID

if test_connection(plc_ams_net_id, ams_port_TC3):
    print("Connection to PLC is ok")
else:
    print("Route needs to be created")
    # create_connection(system_name, plc_ip_address, username, password, ams_port_TC3)


if not check_plc_connection(plc_ams_net_id, ams_port_TC3):
    print("Check your AMS route configuration or TwinCAT installation.")

# If ADS connection is not created, we still can open the connection

net_id = get_local_ams_id()
print(net_id)





# pyads.add_route_to_plc