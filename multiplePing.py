import subprocess
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed

def is_host_reachable(host, timeout=2):
    """Ping the host to check if it is reachable."""
    if platform.system().lower() == "windows":
        ping_cmd = ["ping", "-n", "1", "-w", str(timeout * 1000), host]
    else:
        ping_cmd = ["ping", "-c", "1", "-W", str(timeout), host]

    try:
        subprocess.run(ping_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, timeout=timeout + 1)
        return True
    except subprocess.TimeoutExpired:
        return False
    except subprocess.CalledProcessError:
        return False

import subprocess
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed

def is_host_reachable(host, timeout=2):
    """Ping the host to check if it is reachable."""
    if platform.system().lower() == "windows":
        ping_cmd = ["ping", "-n", "1", "-w", str(timeout * 1000), host]
    else:
        ping_cmd = ["ping", "-c", "1", "-W", str(timeout), host]

    try:
        subprocess.run(ping_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, timeout=timeout + 1)
        return True
    except subprocess.TimeoutExpired:
        return False
    except subprocess.CalledProcessError:
        return False

def find_free_ips(base_ip, max_threads=50):
    """Ping all IPs from .0 to .255 in the given subnet and print the ones that are not reachable (free)."""
    free_ips = []

    def ping_single(ip_suffix):
        ip = f"{base_ip}.{ip_suffix}"
        if not is_host_reachable(ip):
            return ip_suffix  # return just the last part for sorting
        return None

    with ThreadPoolExecutor(max_workers=max_threads) as executor:
        futures = [executor.submit(ping_single, i) for i in range(256)]
        for future in as_completed(futures):
            result = future.result()
            if result is not None:
                free_ips.append(result)

    # Sort numerically and print
    free_ips.sort()
    for suffix in free_ips:
        print(f"{base_ip}.{suffix} is free")

    return [f"{base_ip}.{i}" for i in free_ips]



if __name__ == "__main__":
    subnet = "172.16.12"
    free = find_free_ips(subnet)
    print(f"\nTotal free IPs: {len(free)}")
