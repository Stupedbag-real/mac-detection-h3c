import customtkinter as tk
#from customtkinter import filedialog
from customtkinter import *
import tkinter as ttk
import paramiko
import openpyxl
import sys
import time
import re
import getpass
import io

# Create GUI
class ConsoleOutput(io.StringIO):
    def __init__(self, textbox):
        self.textbox = textbox
        io.StringIO.__init__(self)
    
    def write(self, message):
        self.textbox.configure(state='normal')
        self.textbox.insert('end', message)
        self.textbox.see('end')
        self.textbox.configure(state='disabled')

    def flush(self):
        pass


class Application(tk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MAC address search tool")
        self.geometry("400x400")
        #self.minsize(300, 200)
        self.grid_rowconfigure(0,weight=1)
        self.grid_columnconfigure(0,weight=1)  
        self.grid_rowconfigure(1,weight=1) 
        self.create_widgets()

    #def on_resize(self, event):
        # get the current window size
        #width = self.winfo_width()
        #height = self.winfo_height()

        # check if the size has actually changed
        #if width != self.prev_width or height != self.prev_height:
            #self.prev_width = width
            #self.prev_height = height

        # calculate the new button size
        #button_width = int(width * 0.8)
        #button_height = int(height * 0.1)
        
        # set the new button size
        #self.username_label.configure(width=button_width, height=button_height)
        #self.username_entry.configure(width=button_width, height=button_height)
        #self.password_label.configure(width=button_width, height=button_height)
        #self.password_entry.configure(width=button_width, height=button_height)
        #self.port_label.configure(width=button_width, height=button_height)
        #self.port_entry.configure(width=button_width, height=button_height)
        #self.mac_address_label.configure(width=button_width, height=button_height)
        #self.mac_address_entry.configure(width=button_width, height=button_height)
        #self.file_path_label.configure(width=button_width, height=button_height)
        #self.browse_button.configure(width=button_width, height=button_height)
        #self.search_button.configure(width=button_width, height=button_height)
        #self.quit_button.configure(width=button_width, height=button_height)


    def create_widgets(self):
        self.username_label = tk.CTkLabel(self, text="Enter username:")
        self.username_label.grid(row=0, column=0, sticky="ew")
        self.username_label.grid_propagate(False)
        self.username_entry = tk.CTkEntry(self)
        self.username_entry.grid(row=0, column=1, sticky="ew")
        #self.username_entry.grid_propagate(False)
        self.password_label = tk.CTkLabel(self, text="Enter password:")
        self.password_label.grid(row=1, column=0, sticky="ew")
        self.password_label.grid_propagate(False)
        self.password_entry = tk.CTkEntry(self, show="*")
        self.password_entry.grid(row=1, column=1, sticky="ew")
        #self.password_entry.grid_propagate(False)
        self.port_label = tk.CTkLabel(self, text="Enter physical port to search on:")
        self.port_label.grid(row=2, column=0, sticky="ew")
        self.port_label.grid_propagate(False)
        self.port_entry = tk.CTkEntry(self)
        self.port_entry.grid(row=2, column=1, sticky="ew")
        #self.port_entry.grid_propagate(False)
        self.mac_address_label = tk.CTkLabel(self, text="Enter MAC address to search for:")
        self.mac_address_label.grid(row=3, column=0, sticky="ew")
        self.mac_address_label.grid_propagate(False)
        self.mac_address_entry = tk.CTkEntry(self)
        self.mac_address_entry.grid(row=3, column=1, sticky="ew")
        self.mac_address_entry.grid_propagate(False)
        self.file_path_label = tk.CTkLabel(self, text="Path to file with device IP addresses:")
        self.file_path_label.grid(row=4, column=0, sticky="ew")
        self.file_path_label.grid_propagate(False)
        self.browse_button = tk.CTkButton(self, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=4, column=2, sticky="ew")
        self.browse_button.grid_propagate(False)
        self.browse_entry = CTkEntry(self)
        self.browse_entry.grid(row=4, column=1, sticky="ew")
        self.search_button = tk.CTkButton(self, text="Search", command=self.search)
        self.search_button.grid(row=5, column=0, sticky="nsew")
        self.search_button.grid_propagate(False)
        # Quit button
        self.quit_button = tk.CTkButton(self, text="Quit", command=self.quit)
        self.quit_button.grid(row=5, column=1, sticky="ew")
        self.quit_button.grid_propagate(False)
        self.console_output = tk.CTkTextbox(self, state='disabled')
        self.console_output.grid(row=6, column=0, columnspan=2, padx=20, pady=(20, 0), sticky="nsew")
        sys.stdout = ConsoleOutput(self.console_output)

    def browse_file(self):
        # open file dialog to select file
        file_path_raw = filedialog.askopenfilename()
        self.file_path = file_path_raw
        # set text of entry widget to selected file path
        self.browse_entry.delete(0, tk.END)
        self.browse_entry.insert(0,file_path_raw)
    
    # Create a button to start the search
    def search(self):
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
        while True:
            username = self.username_entry.get()
            if not username:
                print("Please provide a username")
                continue
            break
        while True:
            password = self.password_entry.get()
            if not password:
                print("Please provide a password")
                continue
            break
        # Prompt user for physical port and mac address to search for
        while True:
            port_raw = self.port_entry.get()
            if not port_raw:
                print("Please provide a physical port of device")
                continue
            break
        port = convert_ports(port_raw)
        while True:
            mac_address_raw =self.mac_address_entry.get()
            if not mac_address_raw:
                print("Please provide a mac-address")
                continue
            break
        mac_address = convert_mac_address(mac_address_raw)

        file_path = self.file_path
        if not file_path:
            print("Please provide a path to a file with device IP addresses")
            return

        #file_path = input("Enter path to file with device IP addresses: ")
        #file_path=self.browse_file(file_path)
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
        # Console output
        #self.console = Console(self)
    def quit (self):
        self.destroy()
        
if __name__ == "__main__":
    app=Application()
    app.mainloop()