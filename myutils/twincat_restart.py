"""
Utilities for restarting a TwinCAT runtime via ADS.

Restart sequence
----------------
1. Read initial state.
2. Send RESET.
3. Wait for the state to leave the initial state (confirms RESET was accepted).
4. Wait for the target state (RUN for remote, CONFIG for local).

Known transitions:
  - Remote PLC in RUN:    RUN    -> (transition) -> RUN
  - Local in CONFIG:      CONFIG -> (transition) -> CONFIG
  - Remote PLC in CONFIG: CONFIG -> (transition) -> RUN

"""

import logging
import time
from contextlib import contextmanager

import pyads
from pyads.constants import (
    ADSSTATE_CONFIG,
    ADSSTATE_RESET,
    ADSSTATE_RUN,
)

# Port of the TwinCAT System Service used for system-level state control (RECONFIG, RESET, read_state, etc.).
SYSTEM_SERVICE_PORT = 10000

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Connection helper
# ---------------------------------------------------------------------------

@contextmanager
def _ads_connection(ams_net_id, port=SYSTEM_SERVICE_PORT):
    """Context manager that opens an ADS connection and guarantees close."""
    conn = pyads.Connection(ams_net_id, port)
    conn.open()
    try:
        yield conn
    finally:
        try:
            conn.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# State polling
# ---------------------------------------------------------------------------

def wait_for_ads_state(ams_net_id, target_state, timeout, poll_interval):
    """
    Poll until the ADS state equals *target_state* or *timeout* expires.
    """
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None

    while time.monotonic() < deadline:
        try:
            with _ads_connection(ams_net_id) as conn:
                ads_state, device_state = conn.read_state()

            log.debug("ads_state=%s device_state=%s", ads_state, device_state)

            if ads_state == target_state:
                return ads_state, device_state

        except pyads.ADSError as e:
            last_error = e
            log.debug("Transient ADS error while polling final state: %s", e)

        time.sleep(poll_interval)

    raise TimeoutError(
        f"Timed out waiting for ADS state {target_state}. "
        f"Last ADS error: {last_error}"
    )


def wait_for_state_change(ams_net_id, initial_state, timeout = 30.0, poll_interval = 0.3):
    """
    Poll until the ADS state is different from *initial_state*.

    Used to confirm that a RESET command was actually accepted, the PLC
    must leave its current state before we check for the final target state.

    """
    deadline = time.monotonic() + timeout
    consecutive_errors = 0

    while time.monotonic() < deadline:
        try:
            with _ads_connection(ams_net_id) as conn:
                ads_state, device_state = conn.read_state()

            log.debug("ads_state=%s device_state=%s", ads_state, device_state)

            consecutive_errors = 0

            if ads_state != initial_state:
                log.info("State changed from %s to %s", initial_state, ads_state)
                return ads_state, device_state

        except pyads.ADSError as e:
            consecutive_errors += 1
            log.debug("Transient ADS error (%d consecutive) while polling during wait for state change: %s", consecutive_errors, e)
            if consecutive_errors == 1:
                log.info("Single connection drop detected - assuming restart in progress during wait for state change")
                return None, None
            raise TimeoutError (
                f"Too many consecutive ADS errors - possible network issue"
                f"Last error: {e}"
            )

        time.sleep(poll_interval)

    raise TimeoutError(
        f"Timed out waiting for state to change from {initial_state}. "
    )



# ---------------------------------------------------------------------------
# Core restart logic
# ---------------------------------------------------------------------------

def restart_twincat(ams_net_id, stop_at_config=False, timeout=60, poll_interval=0.3):
    """
    Restart a TwinCAT runtime with ams_net_id
    """
    target_state = ADSSTATE_CONFIG if stop_at_config else ADSSTATE_RUN

    # --- Step 1: read initial state ------------------------------------------
    with _ads_connection(ams_net_id) as conn:
        initial_state, device_state = conn.read_state()
        log.info("Initial state: ads=%s device=%s", initial_state, device_state)

    # --- Step 2: send RESET --------------------------------------------------
    with _ads_connection(ams_net_id) as conn:
        _, device_state = conn.read_state()
        log.info("Sending RESET...")
        conn.write_control(ADSSTATE_RESET, device_state, 0, pyads.PLCTYPE_BYTE)

    # --- Step 3: wait for state to leave initial state -----------------------
    # This confirms RESET was actually accepted, not silently ignored.
    wait_for_state_change(ams_net_id, initial_state)

    # --- Step 4: wait for target state ---------------------------------------
    ads_state, device_state = wait_for_ads_state(
        ams_net_id, target_state, timeout=timeout, poll_interval=poll_interval
    )
    log.info("Reached target state: ads=%s device=%s", ads_state, device_state)

# ---------------------------------------------------------------------------
# Local-machine helpers
# ---------------------------------------------------------------------------

def get_local_ams_netid():
    """
    Return the AMS Net ID of the local TwinCAT installation.
    """
    try:
        pyads.open_port()
        netid = pyads.get_local_address().netid
        log.debug("Local AMS Net ID: %s", netid)
        return netid
    except pyads.ADSError as e:
        raise RuntimeError(
            f"Could not read local AMS Net ID — is TwinCAT running? ({e})"
        ) from e
    except Exception as e:
        raise RuntimeError(f"Unexpected error reading local AMS Net ID: {e}") from e
    finally:
        pyads.close_port()


def restart_local_twincat(timeout=15, poll_interval=0.5):
    """Restart the TwinCAT instance on the local machine."""
    local_netid = get_local_ams_netid()
    restart_twincat(local_netid, stop_at_config=True, timeout=timeout, poll_interval=poll_interval,)



# if __name__ == "__main__":

#     restart_local_twincat()

    # ams_net_id = '10.60.119.126.1.1'
    # restart_twincat(ams_net_id)

    # ams_net_id = '10.60.119.124.1.1'
    # restart_twincat_reset(ams_net_id)
