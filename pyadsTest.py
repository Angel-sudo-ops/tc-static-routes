import pyads
import platform

ams_port_TC3 = pyads.PORT_TC3PLC1    # Default port for TwinCAT 3 PLC
ams_port_TC2 = pyads.PORT_TC2PLC1

username = "Administrator"
password = "1"

system_name = platform.node()

def create_connection(host_name, plc_ip_address, username, password, port):
    ams_net_id = plc_ip_address + ".1.1"

    local_ams_net_id = get_local_ams_id_pyads()

    plc = pyads.Connection(ams_net_id, port)
    plc.open()

    try:
        add_route = pyads.add_route_to_plc(
                local_ams_net_id,
                host_name,
                plc_ip_address,
                username,
                password
            )
    except Exception as e:
        print(f"Exception: {e}")
        add_route = False
    finally:
        if plc.is_open:
            plc.close()
        if add_route:
            print("Route added sucessfully!!")
        else:
            print(":(")



def test_connection(ams_net_id, port):
    plc = pyads.Connection(ams_net_id, port)
    try:
        # Open the connection
        plc.open()

        # If the connection is open, proceed to check the state
        print("PLC connection open")
        state = plc.read_state()
        print(f"PLC Status: {state}")

        # Check if the PLC is in RUN state (state[0] == 5)
        if state[0] == 5:
            print("Connection to PLC established successfully.")
            return True
        else:
            print("PLC is not in RUN state.")
            return False

    except pyads.ADSError as ads_error:
        print(f"ADS Error: {ads_error}")
        return False
    except Exception as e:
        print(f"Unexpected Exception: {e}")
        return False
    finally:
        # Ensure the connection is closed if it was successfully opened
        if plc.is_open:
            plc.close()
   

def get_local_ams_id_pyads_with_instance():
    ams_net_id = None
    try:
        plc = pyads.Connection('0.0.0.0.0.0', pyads.PORT_TC3PLC1)
        plc.open()
        ams_net_id = plc.get_local_address().netid
    except Exception as e:
        print(f"Unexpected error: {e} \nCheck if TwinCAT on local machine is running")
    finally:
        if plc.is_open:
            plc.close()
    return ams_net_id


def get_local_ams_id_pyads():
    ams_net_id=None
    try:
        pyads.open_port()
        ams_net_id = pyads.get_local_address().netid
    except Exception as e:
        print(f"Unexpected error: {e} \nCheck if TwinCAT on local machine is running")
    finally:
        pyads.close_port()
    return ams_net_id

# Define ADS parameters
plc_ip_address = "172.20.2.87"      # Replace with your PLC's IP address
plc_ams_net_id = plc_ip_address + ".1.1" # Replace with your PLC's AMS Net ID

if test_connection(plc_ams_net_id, ams_port_TC2):
    print("Connection to PLC is ok")
else:
    print("Route needs to be created")
    # create_connection(system_name, plc_ip_address, username, password, ams_port_TC3)

# If ADS connection is not created, we still can open the connection

net_id = get_local_ams_id_pyads()
print(net_id)




# pyads.add_route_to_plc