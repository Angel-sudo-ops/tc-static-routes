import paramiko
import threading

ssh_host = "10.77.204.11"  # Your SSH gateway
ssh_user = "Administrator"
ssh_password = "1"
local_port = 5900  # Port you want to listen on locally
remote_host = "192.168.11.6"  # The actual service you're tunneling to
remote_port = 5900  # The remote service port

def forward_tunnel(local_port, remote_host, remote_port, transport):
    """ Opens an SSH tunnel forwarding local_port to remote_host:remote_port """
    try:
        transport.request_port_forward("", local_port)
        while True:
            channel = transport.accept(10)  # Timeout after 10s
            if channel is None:
                continue
            print(f"Tunnel Active: {local_port} -> {remote_host}:{remote_port}")
            threading.Thread(target=handle_connection, args=(channel, remote_host, remote_port)).start()
    except Exception as e:
        print(f"Error in tunnel: {e}")

def handle_connection(channel, remote_host, remote_port):
    """ Handles individual forwarded connections """
    try:
        sock = transport.open_channel("direct-tcpip", (remote_host, remote_port), channel.getpeername())
        if sock is None:
            print("Failed to open channel")
            return
        while True:
            data = channel.recv(1024)
            if not data:
                break
            sock.sendall(data)
            channel.sendall(sock.recv(1024))
    except Exception as e:
        print(f"Error in connection: {e}")
    finally:
        channel.close()

# SSH Connection
try:
    print(f"Connecting to {ssh_host} as {ssh_user}...")
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(ssh_host, port=20022, username=ssh_user, password=ssh_password)

    transport = client.get_transport()
    if transport:
        print("SSH Transport Established")
        tunnel_thread = threading.Thread(target=forward_tunnel, args=(local_port, remote_host, remote_port, transport))
        tunnel_thread.start()
    else:
        print("Failed to establish transport")

except paramiko.AuthenticationException:
    print("Authentication failed, please verify your credentials")
except paramiko.SSHException as ssh_ex:
    print(f"Unable to establish SSH connection: {ssh_ex}")
except Exception as e:
    print(f"Error: {e}")
