import customtkinter as tk
from customtkinter import filedialog
import tkinter
import paramiko
import openpyxl
import sys
import time
# Create GUI
class ConsoleOutput:
    def __init__(self, widget):
        self.widget = widget
    def write(self, message):
        self.widget.insert('end', message)
        self.widget.see('end')


class Application(tk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MAC address search tool")
        self.geometry("1000x800")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        #self.my_frame = ConsoleOutput(self)
        #self.my_frame.grid(row=6, column=0, padx=60, pady=60, sticky="nsew")
        self.create_widgets()


    def create_widgets(self):
        self.console_output = tk.CTkText(self.master, state='disabled')
        self.console_output.pack(expand=True, fill='both')
        sys.stdout = ConsoleOutput(self.console_output)
        # Create labels and entry fields for user input
        self.username_label = tk.CTkLabel(self, text="Enter username:")
        self.username_label.grid(row=0, column=0, sticky="w")
        self.username_entry = tk.CTkEntry(self)
        self.username_entry.grid(row=0, column=1)
        self.password_label = tk.CTkLabel(self, text="Enter password:")
        self.password_label.grid(row=1, column=0, sticky="w")
        self.password_entry = tk.CTkEntry(self, show="*")
        self.password_entry.grid(row=1, column=1)
        self.port_label = tk.CTkLabel(self, text="Enter physical port to search on:")
        self.port_label.grid(row=2, column=0, sticky="w")
        self.port_entry = tk.CTkEntry(self)
        self.port_entry.grid(row=2, column=1)
        self.mac_address_label = tk.CTkLabel(self, text="Enter MAC address to search for:")
        self.mac_address_label.grid(row=3, column=0, sticky="w")
        self.mac_address_entry = tk.CTkEntry(self)
        self.mac_address_entry.grid(row=3, column=1)
        self.file_path_label = tk.CTkLabel(self, text="Path to file with device IP addresses:")
        self.file_path_label.grid(row=4, column=0, sticky="w")
        self.browse_button = tk.CTkButton(self, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=4, column=1, sticky="w")
        self.search_button = tk.CTkButton(self, text="Search", command=self.search)
        self.search_button.grid(row=5, column=1, sticky="w")
    
    def browse_file(self):
        file_path = filedialog.askopenfilename()
        print("Selected file:", file_path)

    # Create a button to start the search
    def search(self):
        # Get user input
        username = self.username_entry.get()
        password = self.password_entry.get()
        port = self.port_entry.get()
        mac_address = self.mac_address_entry.get()
        file_path = filedialog.askopenfilename()

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
        for device in devices:
            try:
                ssh.connect(hostname=device, username=username, password=password, timeout=10)
                print(f"Connected to {device}")
                command = f"display mac-address | i {port}"
                stdin, stdout, stderr = ssh.exec_command(command)
                output = stdout.read().decode('utf-8')
                if mac_address in output:
                    print(f"MAC address {mac_address} found on device {device} on port {port}")
                    command = f"display interface {port}"
                    stdin, stdout, stderr = ssh.exec_command(command)
                    output = stdout.read().decode('utf-8')
                    lines = output.split('\n')
                    for line in lines:
                        if "Description" in line:
                            description = line.split(':')[1].strip()
                            print(f"Port {port} description: {description}")
                    print("\n")
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
        self.search_button = tk.Button(self, text="Search", command=self.search)
        self.search_button.grid(row=5, column=1)
        # Quit button
        self.quit_button = tk.Button(self, text="Quit", command=self.master.quit)
        self.quit_button.grid(row=5, column=1, padx=5, pady=5)
        # Console output
        self.console = Console(self)
if __name__ == "__main__":
    app=Application()
    app.mainloop()