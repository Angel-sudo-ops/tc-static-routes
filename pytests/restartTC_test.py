
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


SYSTEM_SERVICE_PORT = 10000

def wait_for_ads_state(ams_net_id, target_state, timeout=15.0, poll_interval=0.3):
    deadline = time.monotonic() + timeout
    last_error = None

    while time.monotonic() < deadline:
        conn = pyads.Connection(ams_net_id, SYSTEM_SERVICE_PORT)
        try:
            conn.open()
            ads_state, device_state = conn.read_state()
            print(f"[DEBUG] ads_state={ads_state}, device_state={device_state}")

            if ads_state == target_state:
                return ads_state, device_state

        except pyads.ADSError as e:
            last_error = e
            print(f"[DEBUG] read_state failed during transition: {e}")

        finally:
            try:
                conn.close()
            except Exception:
                pass

        time.sleep(poll_interval)

    raise TimeoutError(
        f"Timed out waiting for ADS state {target_state}. "
        f"Last ADS error: {last_error}"
    )


def restart_twincat(ams_net_id, local = False, timeout=15.0, poll_interval=0.3):
    conn = pyads.Connection(ams_net_id, SYSTEM_SERVICE_PORT)

    try:
        conn.open()
        ads_state, device_state = conn.read_state()
        print(f"[INFO] Initial state: {ads_state}")

        if ads_state != ADSSTATE_CONFIG:
            print("[INFO] Sending RECONFIG...")
            conn.write_control(ADSSTATE_RECONFIG, device_state, 0, pyads.PLCTYPE_BYTE)
        else:
            print("[INFO] Already in CONFIG.")

    finally:
        try:
            conn.close()
        except Exception:
            pass

    ads_state, device_state = wait_for_ads_state(
        ams_net_id,
        ADSSTATE_CONFIG,
        timeout=timeout,
        poll_interval=poll_interval
    )
    print(f"[INFO] Reached CONFIG: {ads_state}")

    conn = pyads.Connection(ams_net_id, SYSTEM_SERVICE_PORT)
    try:
        conn.open()
        print("[INFO] Sending RESET...")
        conn.write_control(ADSSTATE_RESET, device_state, 0, pyads.PLCTYPE_BYTE)
    finally:
        try:
            conn.close()
        except Exception:
            pass

    if local:
        target_state = ADSSTATE_CONFIG
    else:
        target_state = ADSSTATE_RUN

    ads_state, device_state = wait_for_ads_state(
        ams_net_id,
        target_state,
        timeout=timeout,
        poll_interval=poll_interval
    )
    print(f"[INFO] Reached target state: {ads_state}")



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
    restart_twincat(local_netid, local=True)



if __name__ == "__main__":

    # restart_local_twincat()

    ams_net_id = '10.60.119.126.1.1'
    restart_twincat(ams_net_id)

# poner en azul el lgv en la tabla hasta que se reinicia 
# rojo cuando manda el reset
# verde cuando ya está en run

# click derecho, y cuando esté haciendo la rutina deselleccionar elemento para poder ver los colores

    