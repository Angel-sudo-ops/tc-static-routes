import subprocess

def restart_twincat_runtime():
    # Stop the TwinCAT runtime service
    subprocess.run(["sc", "stop", "TcSysSrv"], shell=False)
    # Start it again
    subprocess.run(["sc", "start", "TcSysSrv"], shell=False)
