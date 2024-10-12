import asyncio
import struct
import socket
import platform
import pyads
import select

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


def string_to_byte_format(ip_string):
    # Split the string by the dot '.'
    parts = ip_string.split('.')
    
    # Convert each part to an integer and then to a byte
    byte_representation = bytes(int(part) for part in parts)
    
    return byte_representation

def string_to_int_array (ip_string):
    return list(map(int, ip_string.split('.')))

def get_local_ip():
    try:
        # Create a dummy socket to get the local IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0)
        # Connect to a public address (no real connection is made)
        s.connect(('8.8.8.8', 80))
        local_ip = s.getsockname()[0]
        s.close()
        return local_ip
    except Exception as e:
        return f"Error: {e}"



class TcpStateObject:
    def __init__(self):
        self.data = bytearray(1024)  # Buffer size, adjust as needed
        self.CurrentIndex = 0


class RouteManager:
    def __init__(self):
        self.AddRouteSuccess = False
        self.AddRouteError = False
        self._remoteAMSNetID = None
        self.ADSErrorCode = 0
        self.UDPSocket = None
        self.RouteAdded = False

    async def EZRegisterToRemote(self, local_name, local_ip, my_ams_net_id, username, password, remote_ip, use_static_route):
        print("Starting EZRegisterToRemote...")

        my_ip_address = local_ip
        router_table_name = "TCP_" + local_name
        int_send_length = 27 + len(router_table_name) + 15 + len(username) + 5 + len(password) + 5 + len(my_ip_address) + 1

        if not use_static_route:
            int_send_length += 8

        sendbuf = bytearray(int_send_length + 1)

        sendbuf[0] = 3
        sendbuf[1] = 102
        sendbuf[2] = 20
        sendbuf[3] = 113
        sendbuf[4] = 0
        sendbuf[5] = 0
        sendbuf[6] = 0
        sendbuf[7] = 0
        sendbuf[8] = 6
        sendbuf[9] = 0
        sendbuf[10] = 0
        sendbuf[11] = 0

        # Copy AMS Net ID into the buffer at the correct position
        sendbuf[12:18] = my_ams_net_id 

        sendbuf[18] = 16
        sendbuf[19] = 39

        if use_static_route:
            sendbuf[20] = 5
        else:
            sendbuf[20] = 6

        sendbuf[21] = 0
        sendbuf[22] = 0
        sendbuf[23] = 0
        sendbuf[24] = 12
        sendbuf[25] = 0
        sendbuf[26] = len(router_table_name) + 1
        sendbuf[27] = 0
 
        i = 28
        sendbuf[i:i+len(router_table_name)] = router_table_name.encode('ascii')
        i += len(router_table_name)

        sendbuf[i] = 0
        i += 1

        sendbuf[i] = 7
        sendbuf[i+1] = 0
        sendbuf[i+2] = 6
        sendbuf[i+3] = 0
        i += 4

        # Copy AMS Net ID again into the buffer
        sendbuf[i:i+6] = my_ams_net_id
        i += 6

        sendbuf[i] = 13
        sendbuf[i+1] = 0
        i += 2

        sendbuf[i] = len(username) + 1
        sendbuf[i+1] = 0
        i += 2

        sendbuf[i:i+len(username)] = username.encode('ascii')
        i += len(username)

        sendbuf[i] = 0
        i += 1

        sendbuf[i] = 2
        sendbuf[i+1] = 0

        sendbuf[i+2] = len(password) + 1
        sendbuf[i+3] = 0
        i += 4

        sendbuf[i:i+len(password)] = password.encode('ascii')
        i += len(password)

        sendbuf[i] = 0
        i += 1

        sendbuf[i] = 5
        sendbuf[i+1] = 0
        sendbuf[i+2] = len(my_ip_address) + 1
        sendbuf[i+3] = 0
        i += 4

        sendbuf[i:i+len(my_ip_address)] = my_ip_address.encode('ascii')
        i += len(my_ip_address)

        sendbuf[i] = 0
        i += 1

        if len(sendbuf) >= i + 8:
            sendbuf[i:i+8] = struct.pack('BBBBBBBB', 9, 0, 4, 0, 1, 0, 0, 0)

        print(sendbuf)
        print(sendbuf.hex())
        print(len(sendbuf))

        # Now create the UDP socket and send the message asynchronously
        address = (remote_ip, 48899)  # Replace <target_ip> with the actual target IP
        loop = asyncio.get_event_loop()
        self.UDPSocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.UDPSocket.settimeout(5.0)
        # self.UDPSocket.bind((local_ip, 0))
        self.UDPSocket.setblocking(False)
        self.UDPSocket.connect(address)
        # self.UDPSocket.connect(address)
        try:
            retries = 0
            state = TcpStateObject()
            c = 0
            timeout_occurred = False
            while retries < 1 and not self.AddRouteSuccess and not self.AddRouteError:
                # await loop.sock_sendall(self.UDPSocket, sendbuf)
                self.UDPSocket.send(sendbuf)
                print(f"Message sent. Polling for response... (Attempt {retries + 1})")

                # Polling mechanism to check for route addition success or error
                while not self.AddRouteSuccess and not self.AddRouteError and c < 80:
                    # Use select to monitor the socket for readability (timeout of 2 seconds)
                    readable, _, _ = select.select([self.UDPSocket], [], [], 2.0)
                    if readable:
                        await self.DataReceivedA(self.UDPSocket, state)
                    else:
                        print("Timeout while waiting for data")
                        timeout_occurred = True  # Set the timeout flag
                        # break  # Exit the loop if the timeout occurs

                    c += 1

                retries += 1

            # Only set RouteAdded to True if AddRouteSuccess is explicitly True
            if self.AddRouteSuccess:
                self.RouteAdded = True
                print("Route added successfully!")
            elif c >= 40:
                self.RouteAdded = False
                print(f"No response from the remote system after {c} cycles.")
                # raise Exception("No response from remote system. Make sure firewall is off and check username, password, and computer name.")
                print("No response from remote system. Make sure firewall is off and check username, password, and computer name.")
            elif self.AddRouteError:
                self.RouteAdded = False
                print("Error encountered while adding route.")
                # raise Exception("Error setting up remote system, check TwinCATCom for username, password, and computer name.")
                print("Error setting up remote system, check TwinCATCom for username, password, and computer name.")
            elif timeout_occurred:
                self.RouteAdded = False
                print("Route was not added due to select timeout.")

        finally:
            self.UDPSocket.close()
            print("Socket closed.")

    async def DataReceivedA(self, udp_socket, state_obj):
        try:
            # Receive UDP message, block call, wait for data    
            bytes_received = udp_socket.recv(len(state_obj.data) - state_obj.CurrentIndex)

            # Update buffer
            state_obj.CurrentIndex += len(bytes_received)
            state_obj.data[state_obj.CurrentIndex:state_obj.CurrentIndex + len(bytes_received)] = bytes_received

            if state_obj.CurrentIndex > 31:
                AMSNetID = f"{state_obj.data[12]}.{state_obj.data[13]}.{state_obj.data[14]}.{state_obj.data[15]}.{state_obj.data[16]}.{state_obj.data[17]}"
                PortNumber = struct.unpack_from('<H', state_obj.data, 18)[0]
                print(PortNumber)

                self._remoteAMSNetID = AMSNetID
                self.ADSErrorCode = state_obj.data[28] + state_obj.data[29] * 256

                if state_obj.data[27] == 0 and state_obj.data[28] == 0 and state_obj.data[29] == 0 and state_obj.data[30] == 0:
                    self.AddRouteSuccess = True
                    print("SUCCESS!!!!")
                    print(f"Route added successfully. Remote AMSNetID: {self._remoteAMSNetID}")
                else:
                    self.AddRouteError = True
                    self._remoteAMSNetID = "(null AMSID)"
                    print("Route addition failed. Error in response.")
                
                udp_socket.close()
            else:
                # Continue receiving asynchronously until enough data is received
                await self.DataReceivedA(udp_socket, state_obj)
        
        except socket.timeout:
            print("Socket timed out waiting for a response")
        
        except BlockingIOError:
            print("No data available right now, try again later")

        except Exception as e:
            print(f"Error receiving data: {e}")
            return

# Example of using the async RouteManager
async def main():
    route_manager = RouteManager()
    ams_net_id = get_local_ams_netid()
    print(ams_net_id)
    ams_net_id_bit = string_to_byte_format(ams_net_id)
    print(ams_net_id_bit)
    ams_net_id_array = string_to_int_array(ams_net_id)
    print(ams_net_id_array)

    # Remote PLC to create route to
    remote_ip = '10.40.10.74' # LGV05
    user = 'Administrator'
    pass_ = '1'
    
    # Not used
    machine_ip = get_local_ip()
    print(machine_ip)
    
    net_id = ams_net_id.split('.')
    local_ip = '.'.join(net_id[:4])

    system_name = platform.node()
    print(system_name)
    await route_manager.EZRegisterToRemote(system_name, local_ip, ams_net_id_bit, user, pass_, remote_ip, use_static_route=True)

# Run the async main function
asyncio.run(main())
