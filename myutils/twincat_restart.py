"""
Utilities for restarting a TwinCAT runtime via ADS.

Restart sequence
----------------
1. If not already in CONFIG state, send RECONFIG and wait for CONFIG.
2. Send RESET and wait for RUN (or CONFIG when stop_at_config=True).
"""

import logging
import time
from contextlib import contextmanager

import pyads
from pyads.constants import (
    ADSSTATE_CONFIG,
    ADSSTATE_RECONFIG,
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

def wait_for_ads_state(ams_net_id, target_state, timeout=15, poll_interval=0.3):
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
            log.debug("Transient ADS error while polling: %s", e)

        time.sleep(poll_interval)

    raise TimeoutError(
        f"Timed out waiting for ADS state {target_state}. "
        f"Last ADS error: {last_error}"
    )


# ---------------------------------------------------------------------------
# Core restart logic
# ---------------------------------------------------------------------------

def restart_twincat(ams_net_id, stop_at_config=False, timeout=15, poll_interval=0.3):
    """
    Restart a TwinCAT runtime with ams_net_id
    """
    # --- Step 1: move to CONFIG -------------------------------------------
    with _ads_connection(ams_net_id) as conn:
        ads_state, device_state = conn.read_state()
        log.info("Initial state: ads=%s device=%s", ads_state, device_state)

        if ads_state != ADSSTATE_CONFIG:
            log.info("Sending RECONFIG...")
            conn.write_control(ADSSTATE_RECONFIG, device_state, 0, pyads.PLCTYPE_BYTE)
        else:
            log.info("Already in CONFIG, skipping RECONFIG.")

    ads_state, device_state = wait_for_ads_state(
        ams_net_id, ADSSTATE_CONFIG, timeout=timeout, poll_interval=poll_interval
    )
    log.info("Reached CONFIG state.")

    # --- Step 2: reset -------------------------------------------------------
    # Re-read device_state here rather than relying on the value returned by
    # the polling loop, which may reflect the last successful poll some
    # milliseconds ago.
    with _ads_connection(ams_net_id) as conn:
        _, device_state = conn.read_state()
        log.info("Sending RESET...")
        conn.write_control(ADSSTATE_RESET, device_state, 0, pyads.PLCTYPE_BYTE)

    # --- Step 3: wait for final state ----------------------------------------
    target_state = ADSSTATE_CONFIG if stop_at_config else ADSSTATE_RUN
    ads_state, device_state = wait_for_ads_state(
        ams_net_id, target_state, timeout=timeout, poll_interval=poll_interval
    )
    log.info("Reached target state: ads=%s device=%s", ads_state, device_state)
    restart_ok = True

    return restart_ok
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


def restart_local_twincat(stop_at_config=True, timeout=15, poll_interval=0.3):
    """Restart the TwinCAT instance on the local machine."""
    local_netid = get_local_ams_netid()
    restart_twincat(
        local_netid,
        stop_at_config=stop_at_config,
        timeout=timeout,
        poll_interval=poll_interval,
    )
