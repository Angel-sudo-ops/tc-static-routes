import subprocess
import platform
import time
import os

plink_path = r"D:\Documents\PythonPrograms\StaticRoutesCreator\resources\plink.exe"  # update this to your plink location
ssh_host = "172.20.71.61"
ssh_user = "Administrator"
ssh_pass = "1"
ssh_port = 20022

local_port = 5900
remote_ip = "192.168.11.8"
remote_port = 5900

plink_cmd = [
    plink_path,
    "-ssh", f"{ssh_user}@{ssh_host}",
    "-P", str(ssh_port),
    "-pw", ssh_pass,
    "-N",
    "-batch",
    "-L", f"{local_port}:{remote_ip}:{remote_port}"
]

print("Running Plink silently...")

creationflags = subprocess.CREATE_NO_WINDOW if platform.system().lower() == "windows" else 0
process = subprocess.Popen(plink_cmd, creationflags=creationflags)

print(f"Started plink with PID {process.pid}")
print("Tunnel should be open silently. You can RDP/VNC/etc. now.")

try:
    for i in range(60):
        print(f"Tunnel alive... {i+1}/60")
        if process.poll() is not None:
            print("Plink process exited early!")
            break
        time.sleep(1)
finally:
    print("Closing tunnel...")
    process.terminate()
    print("Tunnel closed.")
