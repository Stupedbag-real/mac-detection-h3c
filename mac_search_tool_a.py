#Property of PLcore Petro Lysenko
import getpass
import paramiko
import openpyxl
import time
import re

def convert_ports(port):
    port=re.sub(r"GigabitEthernet", "GE", port_raw)
    print(port)
    return port
def convert_mac_address(mac_address):
    mac_address_st1 = mac_address_raw.lower()
    mac_address_unsub = re.sub(r"[:-]", "", mac_address_st1)
    hex_digits = re.findall(r"[0-9a-fA-F]{2}", mac_address_unsub)
    mac_address = "-".join([f"{hex_digits[i]}{hex_digits[i+1]}" for i in range(0, len(hex_digits), 2)])
    return mac_address

# Prompt user for credentials
while True:
    username = input("Enter username: ")
    if not username:
        print("Please provide a username")
        continue
    break
while True:
    password = getpass.getpass("Enter password: ")
    if not password:
        print("Please provide a username")
        continue
    break
# Prompt user for physical port and mac address to search for
while True:
    port_raw = input("Enter physical port to search on: ")
    if not port_raw:
        print("Please provide a physical port of device")
        continue
    break
port = convert_ports(port_raw)
while True:
    mac_address_raw = input("Enter MAC address to search for: ")
    if not mac_address_raw:
        print("Please provide a mac-address")
        continue
    break
mac_address = convert_mac_address(mac_address_raw)


# Prompt user for path to file with device IP addresses
while True:
    file_path = input("Enter path to file with device IP addresses: ")
    if not file_path:
        print("Please provide a file path")
        continue
    break
#file_path = input("Enter path to file with device IP addresses: ")
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook.active

# Create a list of devices to connect to
devices = []
for row in worksheet.iter_rows(min_row=2, max_col=1):
    device_ip = row[0].value
    devices.append(device_ip)

# Create SSH client object
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# Loop through devices and search for MAC address on specified port
start_time = time.time()
for device in devices:
    try:
        ssh.connect(hostname=device, username=username, password=password, timeout=10)
        print(f"Connected to {device}")
        # Open a shell
        shell = ssh.invoke_shell()
        time.sleep(1)
        commands = [
            f"display mac-address | i {port}",
            f"display interface {port}"
        ]
        # Send commands to the shell
        for command in commands:
            shell.send(command + '\n')
            time.sleep(2)

        # Read the output from the shell
        output = ''
        while shell.recv_ready():
            output += shell.recv(1024).decode('utf-8')
            time.sleep(1)

        if mac_address in output:
            print(f"MAC address {mac_address} found on device {device} on port {port}")
            if "Description" in output:
                description = output.split("Description")[1].split("\n")[0].strip()
                print(f"Port {port} description: {description}")
                ssh.close()
                break
        else:
            print(f"Not found this mac-address")
            print(f"Go next\n⇩")
            #print("\n")
        time.sleep(2)
        ssh.close()
    except paramiko.AuthenticationException:
        print(f"Authentication failed for {device}")
    except paramiko.SSHException:
        print(f"Unable to establish SSH connection to {device}")
    except paramiko.ssh_exception.NoValidConnectionsError:
        print(f"Unable to connect to {device}")
    except Exception as e:
        print(e)
        ssh.close()
end_time = time.time()
print(f"Search completed in {end_time - start_time:.2f} seconds")