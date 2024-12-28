import pyads
from pyads.constants import ADSSTATE_RUN, ADSSTATE_STOP, ADSSTATE_START

def restart_twincat():
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

if __name__ == "__main__":
    restart_twincat()