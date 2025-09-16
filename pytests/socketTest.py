import socket

def is_port_open(host, port, timeout=2):
    """Check if a specific port is open on the host."""
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception as e:
        print(f"Exception at is_port_open for port {port}: {e}")
        return False


host = '10.80.14.11' #'10.49.62.35'
port = '443' # '48898'

result = is_port_open(host, port)
    
print(f"Port {port} for {host} is {'open' if result else 'closed'}")