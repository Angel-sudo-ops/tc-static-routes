
import time
import pyads
from pyads.constants import ADSSTATE_RUN, ADSSTATE_STOP, ADSSTATE_START, ADSSTATE_CONFIG, ADSSTATE_RESET, ADSSTATE_RECONFIG

def restart_twincat_deprecated():
    AMS_NET_ID = "127.0.0.1.1.1"
    PORT = 851
    # Connect to the TwinCAT system
    plc = pyads.Connection(AMS_NET_ID, PORT)  # Replace '127.0.0.1.1.1' with your local AMS Net ID
    try:
        print(f"Connecting to TwinCAT at {AMS_NET_ID}:{PORT}...")
        plc.open()  # Open connection
        print("Connection established.")

        # Read the current ADS state and device state
        # ads_state, device_state = plc.read_state()
        # print(f"Current ADS state: {ads_state}, Device state: {device_state}")
        
        # Switch TwinCAT to STOP state
        print("Switching TwinCAT to STOP state...")
        plc.write_control(ADSSTATE_STOP, 4, 0, pyads.PLCTYPE_BYTE)
        print("TwinCAT is now in STOP state.")

        # Switch TwinCAT back to RUN state
        print("Switching TwinCAT to RUN state...")
        plc.write_control(ADSSTATE_RUN, 5, 0, pyads.PLCTYPE_BYTE)
        print("TwinCAT is now in RUN state.")

    except pyads.ADSError as e:
        print(f"ADS Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        plc.close()



def restart_twincat(net_id, timeout=50.0, poll_interval=0.1):

    system_service_port = 10000
    plc = pyads.Connection(net_id, system_service_port)

    try:
        plc.open()

        ads_state, device_state = plc.read_state()
        print(f"[INFO] Current ADS state: {ads_state}.")

        # If already in CONFIG, skip RECONFIG
        if ads_state != ADSSTATE_CONFIG:
            print("[INFO] Sending RECONFIG...")
            plc.write_control(
                ADSSTATE_RECONFIG,
                device_state,
                0,
                pyads.PLCTYPE_BYTE
            )

            ads_state, device_state = wait_for_ads_state(
                plc,
                ADSSTATE_CONFIG,
                timeout,
                poll_interval
            )
            print((f"[INFO] Reached CONFIG state"))
        
        # Send RESET to return to RUN
        print("[INFO] Sending RESET...")
        plc.write_control(
            ADSSTATE_RESET, 
            device_state, 
            0, 
            pyads.PLCTYPE_BYTE
        )

        ads_state, device_state = wait_for_ads_state(
            plc, 
            ADSSTATE_RUN,
            timeout,
            poll_interval
        )
        print(f"[SUCCESS] TwinCAT is now in RUN state")

    except TimeoutError as te:
        print(f"[TIMEOUT {te}")
    # except Exception as e:
    #     print(f"[ERROR] {e}")

    finally:
        plc.close()
        print("[INFO] Connection closed")


def wait_for_ads_state(conn, target_state, timeout=15.0, poll_interval=0.1):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        ads_state, device_state = conn.read_state()
        if ads_state == target_state:
            print("state reached")
            return ads_state, device_state
        time.sleep(poll_interval)
        print("waiting for state")
    raise TimeoutError(f"Timed out waiting for ADS state {target_state}")



def get_local_ams_netid():
    ams_net_id=None
    try:
        pyads.open_port()
        ams_net_id = pyads.get_local_address().netid
        print (ams_net_id)
    except Exception as e:
        print(f"Unexpected error: {e} \nCheck if TwinCAT on local machine is running")
    finally:
        pyads.close_port()

    return ams_net_id

def restart_local_twincat():
    local_netid = get_local_ams_netid()
    restart_twincat(local_netid)



if __name__ == "__main__":

    # restart_local_twincat()

    ams_net_id = '10.60.119.126.1.1'
    restart_twincat(ams_net_id)
    