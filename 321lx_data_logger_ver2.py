import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import threading
import time
import datetime
import re
import os
import json
import platform
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import socket

class BalanceLogger:
    TEMP_FILE = "balance_data.tmp.json"

    def __init__(self, root):
        self.root = root
        self.root.title("321 LX Scale Data Logger")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)

        # Presets and email settings
        self.presets_file = "scale_presets.json"
        self.presets = self.load_presets()
        self.email_settings_file = "email_settings.json"
        self.email_settings = self.load_email_settings()

        # Serial parameters
        self.serial_params = {
            'port': tk.StringVar(),
            'baudrate': tk.IntVar(value=9600),
            'bytesize': tk.IntVar(value=serial.SEVENBITS),
            'parity': tk.StringVar(value='ODD'),
            'stopbits': tk.IntVar(value=serial.STOPBITS_ONE),
            'flowcontrol': tk.StringVar(value='NONE')
        }
        self.connection_type = tk.StringVar(value='Serial')  # 'Serial' or 'Ethernet'
        self.tcp_ip = tk.StringVar(value='192.168.1.100')
        self.tcp_port = tk.IntVar(value=8000)

        # Email parameters
        self.email_params = {
            'smtp_server': tk.StringVar(value=self.email_settings.get('smtp_server', 'smtp.gmail.com')),
            'smtp_port': tk.IntVar(value=self.email_settings.get('smtp_port', 587)),
            'username': tk.StringVar(value=self.email_settings.get('username', '')),
            'password': tk.StringVar(value=self.email_settings.get('password', '')),
            'sender': tk.StringVar(value=self.email_settings.get('sender', '')),
            'default_recipient': tk.StringVar(value=self.email_settings.get('default_recipient', '')),
            'default_subject': tk.StringVar(value=self.email_settings.get('default_subject', 'Weight Data Export')),
            'default_message': tk.StringVar(value=self.email_settings.get('default_message', 'Please find attached the weight data export.'))
        }

        # App state
        self.serial_port = None
        self.tcp_socket = None
        self.data = []
        self.file_path = tk.StringVar(value="balance_data.csv")
        self.sample_counter = 1
        self.read_thread = None
        self.device_name = tk.StringVar(value="")

        # UI
        self.create_ui()
        self.refresh_ports()
        self.create_help_messages()
        self.root.bind("<Configure>", self.on_window_resize)
        self.load_temp_data()

    # --- Crash Recovery ---
    def save_temp_data(self):
        try:
            with open(self.TEMP_FILE, "w") as f:
                json.dump({
                    "data": self.data,
                    "sample_counter": self.sample_counter,
                    "device_name": self.device_name.get()
                }, f)
        except Exception as e:
            print(f"Failed to save temp data: {e}")

    def load_temp_data(self):
        if os.path.exists(self.TEMP_FILE):
            try:
                with open(self.TEMP_FILE, "r") as f:
                    temp = json.load(f)
                if messagebox.askyesno("Restore Session", "Restore previous unsaved session?"):
                    self.data = temp.get("data", [])
                    self.sample_counter = temp.get("sample_counter", 1)
                    self.device_name.set(temp.get("device_name", ""))
                    self.refresh_table()
                    self.show_status("Session restored from crash recovery.", color="blue")
                else:
                    self.clear_temp_data()
            except Exception as e:
                print(f"Failed to load temp data: {e}")

    def clear_temp_data(self):
        try:
            if os.path.exists(self.TEMP_FILE):
                os.remove(self.TEMP_FILE)
        except Exception as e:
            print(f"Failed to delete temp data: {e}")

    # --- UI Construction ---
    def create_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        main_tab = ttk.Frame(self.notebook)
        self.notebook.add(main_tab, text="Data Logger")
        settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(settings_tab, text="Settings")

        self.create_main_tab_ui(main_tab)
        self.create_settings_tab_ui(settings_tab)

    def create_main_tab_ui(self, parent):
        # --- Preset selection at top ---
        presets_frame = ttk.Frame(parent)
        presets_frame.pack(fill=tk.X, pady=3)
        ttk.Label(presets_frame, text="Scale Preset:").pack(side=tk.LEFT, padx=5)
        self.presets_combo = ttk.Combobox(presets_frame, values=list(self.presets.keys()), state="readonly")
        self.presets_combo.pack(side=tk.LEFT, padx=5)
        self.presets_combo.bind('<<ComboboxSelected>>', self.apply_preset)
        ttk.Button(presets_frame, text="+", command=self.save_new_preset, width=2).pack(side=tk.LEFT, padx=5)
        self.add_help_button(presets_frame, 'new_preset').pack(side=tk.LEFT, padx=5)
        self.presets_menu = tk.Menu(self.root, tearoff=0)
        self.presets_menu.add_command(label="Delete Preset", command=self.delete_preset)
        self.presets_combo.bind("<Button-3>", self.show_presets_menu)

        # --- Connection settings: vertical stack ---
        conn_frame = ttk.LabelFrame(parent, text="Connection Settings", padding="10")
        conn_frame.pack(fill=tk.X, pady=3)
        # Connection type
        conn_type_frame = ttk.Frame(conn_frame)
        conn_type_frame.pack(fill=tk.X, pady=2)
        ttk.Label(conn_type_frame, text="Connection Type:").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(conn_type_frame, text="Serial", variable=self.connection_type, value="Serial", command=self.toggle_conn_type).pack(side=tk.LEFT)
        ttk.Radiobutton(conn_type_frame, text="Ethernet (RJ45)", variable=self.connection_type, value="Ethernet", command=self.toggle_conn_type).pack(side=tk.LEFT)

        # Serial settings
        self.serial_settings_frame = ttk.Frame(conn_frame)
        self.serial_settings_frame.pack(fill=tk.X, pady=2)
        ttk.Label(self.serial_settings_frame, text="COM Port:").grid(row=0, column=0, padx=5, sticky='e')
        self.port_cb = ttk.Combobox(self.serial_settings_frame, width=15, textvariable=self.serial_params['port'])
        self.port_cb.grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Button(self.serial_settings_frame, text="↻", command=self.refresh_ports, width=2).grid(row=0, column=2, sticky='w')
        self.add_help_button(self.serial_settings_frame, 'com_port').grid(row=0, column=3, padx=5, sticky='w')

        ttk.Label(self.serial_settings_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, sticky='e')
        self.baudrate_combo = ttk.Combobox(self.serial_settings_frame, textvariable=self.serial_params['baudrate'], values=[1200, 2400, 4800, 9600, 19200, 38400, 57600], width=10)
        self.baudrate_combo.grid(row=1, column=1, padx=5, sticky='ew')
        self.add_help_button(self.serial_settings_frame, 'baud_rate').grid(row=1, column=2, padx=5, sticky='w')

        ttk.Label(self.serial_settings_frame, text="Data Bits:").grid(row=2, column=0, padx=5, sticky='e')
        self.databits_combo = ttk.Combobox(self.serial_settings_frame, textvariable=self.serial_params['bytesize'], values=[7, 8], state="readonly", width=5)
        self.databits_combo.grid(row=2, column=1, padx=5, sticky='ew')
        self.add_help_button(self.serial_settings_frame, 'data_bits').grid(row=2, column=2, padx=5, sticky='w')

        ttk.Label(self.serial_settings_frame, text="Parity:").grid(row=3, column=0, padx=5, sticky='e')
        self.parity_combo = ttk.Combobox(self.serial_settings_frame, textvariable=self.serial_params['parity'], values=['NONE', 'EVEN', 'ODD', 'MARK', 'SPACE'], width=8)
        self.parity_combo.grid(row=3, column=1, padx=5, sticky='ew')
        self.add_help_button(self.serial_settings_frame, 'parity').grid(row=3, column=2, padx=5, sticky='w')

        ttk.Label(self.serial_settings_frame, text="Stop Bits:").grid(row=4, column=0, padx=5, sticky='e')
        self.stopbits_combo = ttk.Combobox(self.serial_settings_frame, textvariable=self.serial_params['stopbits'], values=[1, 2], width=5)
        self.stopbits_combo.grid(row=4, column=1, padx=5, sticky='ew')

        ttk.Label(self.serial_settings_frame, text="Flow Control:").grid(row=5, column=0, padx=5, sticky='e')
        self.flowcontrol_combo = ttk.Combobox(self.serial_settings_frame, textvariable=self.serial_params['flowcontrol'], values=['NONE', 'XON/XOFF', 'HARDWARE'], width=10)
        self.flowcontrol_combo.grid(row=5, column=1, padx=5, sticky='ew')
        self.add_help_button(self.serial_settings_frame, 'flow_control').grid(row=5, column=2, padx=5, sticky='w')

        # Serial connect button
        self.connect_btn = ttk.Button(self.serial_settings_frame, text="Connect", command=self.toggle_connection)
        self.connect_btn.grid(row=6, column=0, columnspan=2, pady=5)
        self.add_help_button(self.serial_settings_frame, 'connection').grid(row=6, column=2, padx=5, sticky='w')

        # Ethernet settings
        self.ethernet_settings_frame = ttk.Frame(conn_frame)
        # Hidden by default
        self.ethernet_settings_frame.pack_forget()
        ttk.Label(self.ethernet_settings_frame, text="IP Address:").grid(row=0, column=0, padx=5, sticky='e')
        ip_entry = ttk.Entry(self.ethernet_settings_frame, textvariable=self.tcp_ip, width=15)
        ip_entry.grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Label(self.ethernet_settings_frame, text="Port:").grid(row=1, column=0, padx=5, sticky='e')
        port_entry = ttk.Entry(self.ethernet_settings_frame, textvariable=self.tcp_port, width=8)
        port_entry.grid(row=1, column=1, padx=5, sticky='ew')
        self.eth_connect_btn = ttk.Button(self.ethernet_settings_frame, text="Connect", command=self.toggle_connection)
        self.eth_connect_btn.grid(row=2, column=0, columnspan=2, pady=5)

        # Status indicator
        status_frame = ttk.Frame(conn_frame)
        status_frame.pack(fill=tk.X, pady=2)
        self.status_canvas = tk.Canvas(status_frame, width=20, height=20)
        self.status_canvas.pack(side=tk.LEFT)
        self.status_indicator = self.status_canvas.create_oval(2, 2, 18, 18, fill='grey')
        self.status_label = ttk.Label(status_frame, text="Not Connected")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # --- Data management controls ---
        log_frame = ttk.LabelFrame(parent, text="Data Management", padding="10")
        log_frame.pack(fill=tk.X, pady=3)
        ttk.Label(log_frame, text="Export to:").pack(side=tk.LEFT)
        ttk.Entry(log_frame, textvariable=self.file_path, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(log_frame, text="Browse...", command=self.browse_file).pack(side=tk.LEFT)
        ttk.Button(log_frame, text="Export Results", command=self.save_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_frame, text="Open File", command=self.open_exported_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_frame, text="Email Data", command=self.send_email).pack(side=tk.LEFT, padx=5)
        self.add_help_button(log_frame, 'email').pack(side=tk.LEFT, padx=5)
        ttk.Button(log_frame, text="Reset", command=self.reset_data).pack(side=tk.LEFT, padx=5)

        # --- Data table and display (expands with window) ---
        table_frame = ttk.LabelFrame(parent, text="Weight Data", padding="10")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=3)
        self.current_weight = ttk.Label(table_frame, text="0.000 g", font=("Arial", 24))
        self.current_weight.pack(pady=5)

        columns = ("sample_name", "weight", "unit", "device", "comments")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col, width in zip(columns, [130, 100, 60, 160, 200]):
            self.tree.heading(col, text=col.replace("_", " ").title())
            self.tree.column(col, width=width, anchor='center')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind('<Double-1>', self.on_treeview_double_click)

        # Status bar for visual feedback
        self.statusbar = ttk.Label(parent, text="", relief=tk.SUNKEN, anchor="w")
        self.statusbar.pack(fill=tk.X, side=tk.BOTTOM)

    def toggle_conn_type(self):
        if self.connection_type.get() == "Serial":
            self.ethernet_settings_frame.pack_forget()
            self.serial_settings_frame.pack(fill=tk.X, pady=2)
        else:
            self.serial_settings_frame.pack_forget()
            self.ethernet_settings_frame.pack(fill=tk.X, pady=2)

    def show_status(self, msg, color="black"):
        self.statusbar.config(text=msg, foreground=color)
        self.root.after(3000, lambda: self.statusbar.config(text="", foreground="black"))

    def reset_data(self):
        if messagebox.askyesno("Reset All Data", "Are you sure you want to clear all data and reset the sample counter?"):
            self.data.clear()
            self.sample_counter = 1
            self.device_name.set("")
            for i in self.tree.get_children():
                self.tree.delete(i)
            self.current_weight.config(text="0.000 g")
            self.clear_temp_data()
            self.show_status("All data cleared and sample counter reset.", color="red")

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for row in self.data:
            values = (
                row["sample_name"],
                f"{row['weight']:.3f}",
                row["unit"],
                row.get("device", ""),
                row["comments"]
            )
            self.tree.insert("", "end", values=values)
        if self.data:
            self.current_weight.config(text=f"{self.data[-1]['weight']:.3f} {self.data[-1]['unit']}")
        else:
            self.current_weight.config(text="0.000 g")

    def refresh_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_cb['values'] = ports
        if ports:
            self.port_cb.set(ports[0])
        else:
            self.port_cb.set("")

    def toggle_connection(self):
        if self.connection_type.get() == "Serial":
            if self.serial_port is None:
                try:
                    parity_map = {
                        'NONE': serial.PARITY_NONE,
                        'EVEN': serial.PARITY_EVEN,
                        'ODD': serial.PARITY_ODD,
                        'MARK': serial.PARITY_MARK,
                        'SPACE': serial.PARITY_SPACE
                    }
                    flowcontrol_map = {
                        'NONE': {'xonxoff': False, 'rtscts': False},
                        'XON/XOFF': {'xonxoff': True, 'rtscts': False},
                        'HARDWARE': {'xonxoff': False, 'rtscts': True}
                    }
                    self.serial_port = serial.Serial(
                        port=self.port_cb.get(),
                        baudrate=self.serial_params['baudrate'].get(),
                        bytesize=self.serial_params['bytesize'].get(),
                        parity=parity_map[self.serial_params['parity'].get()],
                        stopbits=self.serial_params['stopbits'].get(),
                        **flowcontrol_map[self.serial_params['flowcontrol'].get()],
                        timeout=2
                    )
                    self.connect_btn.config(text="Disconnect")
                    self.update_status("Connected", 'green')
                    self.read_thread = threading.Thread(target=self.read_serial, daemon=True)
                    self.read_thread.start()
                except Exception as e:
                    self.update_status(f"Error: {str(e)}", 'red')
                    messagebox.showerror("Connection Error", str(e))
            else:
                self.serial_port.close()
                self.serial_port = None
                self.connect_btn.config(text="Connect")
                self.update_status("Not Connected", 'grey')
        else:
            if self.tcp_socket is None:
                try:
                    self.tcp_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    self.tcp_socket.settimeout(5)
                    self.tcp_socket.connect((self.tcp_ip.get(), self.tcp_port.get()))
                    self.eth_connect_btn.config(text="Disconnect")
                    self.update_status("Connected (Ethernet)", 'green')
                    self.read_thread = threading.Thread(target=self.read_tcp, daemon=True)
                    self.read_thread.start()
                except Exception as e:
                    self.update_status(f"Error: {str(e)}", 'red')
                    messagebox.showerror("Ethernet Error", str(e))
                    self.tcp_socket = None
            else:
                self.tcp_socket.close()
                self.tcp_socket = None
                self.eth_connect_btn.config(text="Connect")
                self.update_status("Not Connected", 'grey')

    def update_status(self, text, color='grey'):
        self.status_canvas.itemconfig(self.status_indicator, fill=color)
        self.status_label.config(text=text)
        self.root.update_idletasks()
        self.show_status(text, color=color)

    def read_serial(self):
        buffer = ""
        while self.serial_port and self.serial_port.is_open:
            try:
                bytes_to_read = self.serial_port.in_waiting
                if bytes_to_read > 0:
                    raw_data = self.serial_port.read(bytes_to_read)
                    buffer += raw_data.decode('ascii', errors='replace')
                    while '\r\n' in buffer:
                        line, buffer = buffer.split('\r\n', 1)
                        self.process_data(line.strip())
                        self.save_temp_data()
                time.sleep(0.1)
            except serial.SerialException:
                self.update_status("Connection Lost", 'red')
                break
            except Exception as e:
                print(f"Serial error: {str(e)}")
                break

    def read_tcp(self):
        buffer = ""
        while self.tcp_socket:
            try:
                data = self.tcp_socket.recv(1024)
                if not data:
                    break
                buffer += data.decode('ascii', errors='replace')
                while '\r\n' in buffer:
                    line, buffer = buffer.split('\r\n', 1)
                    self.process_data(line.strip())
                    self.save_temp_data()
                time.sleep(0.1)
            except socket.error as e:
                self.update_status(f"TCP Error: {e}", 'red')
                break
            except Exception as e:
                print(f"TCP error: {str(e)}")
                break

    def process_data(self, data):
        try:
            clean_data = re.sub(r'[^\d\+-\.gk]', '', data)
            if not clean_data:
                return
            weight_match = re.search(r'([+-]?\d+\.\d+)', clean_data)
            if not weight_match:
                return
            weight = float(weight_match.group(1))
            unit = 'g' if 'g' in clean_data.lower() else 'kg' if 'kg' in clean_data.lower() else '?'
            sample_name = f"Sample_{self.sample_counter}"
            self.sample_counter += 1
            device = self.presets_combo.get() or self.device_name.get() or "Unknown"
            self.device_name.set(device)
            values = (sample_name, f"{weight:.3f}", unit, device, "")
            iid = self.tree.insert("", "end", values=values)
            self.data.append({
                "sample_name": sample_name,
                "weight": weight,
                "unit": unit,
                "device": device,
                "comments": "",
                "iid": iid
            })
            self.tree.yview_moveto(1)
            self.current_weight.config(text=f"{weight:.3f} {unit}")
            self.save_temp_data()
        except Exception as e:
            print(f"Data processing error: {str(e)}")

    def on_treeview_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        rowid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        col_idx = int(col.replace('#', '')) - 1
        # Only allow editing Sample Name (col 0) and Comments (col 4)
        if col_idx not in [0, 4]:
            return
        x, y, width, height = self.tree.bbox(rowid, col)
        value = self.tree.set(rowid, column=self.tree["columns"][col_idx])
        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value)
        entry.focus()
        def save_edit(event=None):
            new_value = entry.get()
            self.tree.set(rowid, column=self.tree["columns"][col_idx], value=new_value)
            for item in self.data:
                if item["iid"] == rowid:
                    if col_idx == 0:
                        item["sample_name"] = new_value
                    else:
                        item["comments"] = new_value
                    break
            entry.destroy()
            self.save_temp_data()
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def save_data(self):
        if not self.data:
            messagebox.showinfo("Info", "No data to export")
            return
        try:
            with open(self.file_path.get(), 'a', newline='') as f:
                writer = csv.writer(f)
                export_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S %Z")
                writer.writerow([f"Exported: {export_time}"])
                writer.writerow(["Sample Name", "Weight", "Units", "Device", "Comments"])
                for row in self.data:
                    writer.writerow([
                        row["sample_name"],
                        row["weight"],
                        row["unit"],
                        row.get("device", ""),
                        row["comments"]
                    ])
            messagebox.showinfo("Success", f"Data exported to {self.file_path.get()}")
            self.show_status("Data exported.", color="green")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data: {str(e)}")

    def open_exported_file(self):
        file_path = self.file_path.get()
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File not found: {file_path}")
            return
        try:
            if platform.system() == 'Windows':
                os.startfile(file_path)
            elif platform.system() == 'Darwin':
                subprocess.call(['open', file_path])
            else:
                subprocess.call(['xdg-open', file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")

    def add_help_button(self, parent, help_key):
        return ttk.Button(parent, text="?", width=2, command=lambda: messagebox.showinfo("Help", self.help_texts[help_key]))

    def load_presets(self):
        try:
            with open(self.presets_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            default_presets = {
                "Precisa 321 LX": {
                    "baudrate": 9600,
                    "bytesize": 7,
                    "parity": "ODD",
                    "stopbits": 1,
                    "flowcontrol": "XON/XOFF"
                }
            }
            with open(self.presets_file, 'w') as f:
                json.dump(default_presets, f, indent=2)
            return default_presets
        except Exception as e:
            messagebox.showerror("Preset Error", f"Failed to load presets: {str(e)}")
            return {}

    def apply_preset(self, event=None):
        preset_name = self.presets_combo.get()
        if preset_name in self.presets:
            preset = self.presets[preset_name]
            try:
                self.serial_params['baudrate'].set(int(preset['baudrate']))
                self.serial_params['bytesize'].set(int(preset['bytesize']))
                self.serial_params['parity'].set(preset['parity'])
                self.serial_params['stopbits'].set(int(preset['stopbits']))
                self.serial_params['flowcontrol'].set(preset['flowcontrol'])
                self.device_name.set(preset_name)
            except KeyError as e:
                messagebox.showerror("Preset Error", f"Missing key in preset: {str(e)}")

    def save_new_preset(self):
        preset_dialog = tk.Toplevel(self.root)
        preset_dialog.title("New Preset Configuration")
        preset_dialog.grab_set()
        entries = {}
        ttk.Label(preset_dialog, text="Preset Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        name_entry = ttk.Entry(preset_dialog)
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        params = [
            ("Baud Rate:", 'baudrate', [str(x) for x in [1200, 2400, 4800, 9600, 19200, 38400, 57600]]),
            ("Data Bits:", 'bytesize', ['7', '8']),
            ("Parity:", 'parity', ['NONE', 'EVEN', 'ODD', 'MARK', 'SPACE']),
            ("Stop Bits:", 'stopbits', ['1', '2']),
            ("Flow Control:", 'flowcontrol', ['NONE', 'XON/XOFF', 'HARDWARE'])
        ]
        for row, (label, param, values) in enumerate(params, 1):
            ttk.Label(preset_dialog, text=label).grid(row=row, column=0, padx=5, pady=2, sticky='w')
            combo = ttk.Combobox(preset_dialog, values=values)
            combo.set(str(self.serial_params[param].get()))
            combo.grid(row=row, column=1, padx=5, pady=2, sticky='ew')
            entries[param] = combo
        preset_dialog.columnconfigure(1, weight=1)
        def save_preset():
            preset_name = name_entry.get()
            if not preset_name:
                messagebox.showerror("Error", "Preset name cannot be empty")
                return
            try:
                new_preset = {
                    'baudrate': int(entries['baudrate'].get()),
                    'bytesize': int(entries['bytesize'].get()),
                    'parity': entries['parity'].get(),
                    'stopbits': int(entries['stopbits'].get()),
                    'flowcontrol': entries['flowcontrol'].get()
                }
            except ValueError:
                messagebox.showerror("Error", "Invalid numeric values")
                return
            self.presets[preset_name] = new_preset
            try:
                with open(self.presets_file, 'w') as f:
                    json.dump(self.presets, f, indent=2)
                self.presets_combo['values'] = list(self.presets.keys())
                preset_dialog.destroy()
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save preset: {str(e)}")
        ttk.Button(preset_dialog, text="Save", command=save_preset).grid(row=6, columnspan=2, pady=10)

    def delete_preset(self):
        selected = self.presets_combo.get()
        if selected and selected in self.presets:
            del self.presets[selected]
            try:
                with open(self.presets_file, 'w') as f:
                    json.dump(self.presets, f, indent=2)
                self.presets_combo['values'] = list(self.presets.keys())
                self.presets_combo.set('')
            except Exception as e:
                messagebox.showerror("Delete Error", f"Failed to delete preset: {str(e)}")

    def show_presets_menu(self, event):
        self.presets_menu.post(event.x_root, event.y_root)

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)

    # --- Email and Settings ---
    def create_settings_tab_ui(self, parent):
        email_frame = ttk.LabelFrame(parent, text="Email Settings", padding="10")
        email_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        email_frame.columnconfigure(1, weight=1)
        ttk.Label(email_frame, text="SMTP Server:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(email_frame, textvariable=self.email_params['smtp_server']).grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="SMTP Port:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(email_frame, textvariable=self.email_params['smtp_port']).grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="Username:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(email_frame, textvariable=self.email_params['username']).grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="Password:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        password_entry = ttk.Entry(email_frame, textvariable=self.email_params['password'], show="*")
        password_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="Default Sender:").grid(row=4, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(email_frame, textvariable=self.email_params['sender']).grid(row=4, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="Default Recipient:").grid(row=5, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(email_frame, textvariable=self.email_params['default_recipient']).grid(row=5, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="Default Subject:").grid(row=6, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(email_frame, textvariable=self.email_params['default_subject']).grid(row=6, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(email_frame, text="Default Message:").grid(row=7, column=0, padx=5, pady=5, sticky='w')
        message_frame = ttk.Frame(email_frame)
        message_frame.grid(row=7, column=1, padx=5, pady=5, sticky='ew')
        message_text = tk.Text(message_frame, height=4, width=40)
        message_text.pack(fill=tk.BOTH, expand=True)
        message_text.insert("1.0", self.email_params['default_message'].get())
        def update_message_var():
            self.email_params['default_message'].set(message_text.get("1.0", "end-1c"))
        button_frame = ttk.Frame(email_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=10)
        self.add_help_button(button_frame, 'smtp_settings').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Test Connection", command=self.test_email_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save Settings", command=lambda: [update_message_var(), self.save_email_settings()]).pack(side=tk.LEFT, padx=5)
        note_frame = ttk.LabelFrame(parent, text="Note", padding="10")
        note_frame.pack(fill=tk.X, padx=10, pady=10)
        note_text = (
            "For Gmail and many other providers, you'll need to use an App Password instead of your regular password.\n"
            "For Gmail:\n"
            "1. Enable 2-Step Verification in your Google Account\n"
            "2. Create an App Password (Google Account → Security → App Passwords)\n"
            "3. Use that password here instead of your regular password"
        )
        ttk.Label(note_frame, text=note_text, wraplength=600, justify="left").pack(fill=tk.X)

    def test_email_connection(self):
        try:
            server = smtplib.SMTP(
                self.email_params['smtp_server'].get(),
                self.email_params['smtp_port'].get()
            )
            server.starttls()
            server.login(
                self.email_params['username'].get(),
                self.email_params['password'].get()
            )
            server.quit()
            messagebox.showinfo("Success", "SMTP connection successful!")
        except Exception as e:
            messagebox.showerror("Connection Error", f"Failed to connect: {str(e)}")

    def send_email(self):
        if not self.data:
            messagebox.showerror("Error", "No data to send. Please log some weight measurements first.")
            return
        self.save_data()
        email_dialog = tk.Toplevel(self.root)
        email_dialog.title("Send Data via Email")
        email_dialog.grab_set()
        ttk.Label(email_dialog, text="To:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        recipient = ttk.Entry(email_dialog, width=30)
        recipient.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        recipient.insert(0, self.email_params['default_recipient'].get())
        ttk.Label(email_dialog, text="Subject:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        subject = ttk.Entry(email_dialog, width=30)
        subject.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        subject.insert(0, self.email_params['default_subject'].get())
        ttk.Label(email_dialog, text="Message:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        message = tk.Text(email_dialog, width=30, height=5)
        message.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        message.insert("1.0", self.email_params['default_message'].get())
        email_dialog.columnconfigure(1, weight=1)
        def send():
            try:
                msg = MIMEMultipart()
                msg['Subject'] = subject.get()
                msg['From'] = self.email_params['sender'].get()
                msg['To'] = recipient.get()
                msg.attach(MIMEText(message.get("1.0", tk.END), 'plain'))
                file_path = self.file_path.get()
                with open(file_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
                    msg.attach(part)
                server = smtplib.SMTP(
                    self.email_params['smtp_server'].get(),
                    self.email_params['smtp_port'].get()
                )
                server.starttls()
                server.login(
                    self.email_params['username'].get(),
                    self.email_params['password'].get()
                )
                server.sendmail(
                    self.email_params['sender'].get(),
                    recipient.get(),
                    msg.as_string()
                )
                server.quit()
                messagebox.showinfo("Success", "Email sent successfully!")
                email_dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email: {str(e)}")
        ttk.Button(email_dialog, text="Send", command=send).grid(row=3, column=1, pady=10)
        ttk.Button(email_dialog, text="Cancel", command=email_dialog.destroy).grid(row=3, column=0, pady=10)

    def load_email_settings(self):
        try:
            with open(self.email_settings_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            default_settings = {
                'smtp_server': 'smtp.gmail.com',
                'smtp_port': 587,
                'username': '',
                'password': '',
                'sender': '',
                'default_recipient': '',
                'default_subject': 'Weight Data Export',
                'default_message': 'Please find attached the weight data export.'
            }
            with open(self.email_settings_file, 'w') as f:
                json.dump(default_settings, f, indent=2)
            return default_settings
        except Exception as e:
            messagebox.showerror("Settings Error", f"Failed to load email settings: {str(e)}")
            return {}

    def save_email_settings(self):
        try:
            settings = {
                'smtp_server': self.email_params['smtp_server'].get(),
                'smtp_port': self.email_params['smtp_port'].get(),
                'username': self.email_params['username'].get(),
                'password': self.email_params['password'].get(),
                'sender': self.email_params['sender'].get(),
                'default_recipient': self.email_params['default_recipient'].get(),
                'default_subject': self.email_params['default_subject'].get(),
                'default_message': self.email_params['default_message'].get()
            }
            with open(self.email_settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
            messagebox.showinfo("Success", "Email settings saved successfully")
        except Exception as e:
            messagebox.showerror("Settings Error", f"Failed to save email settings: {str(e)}")

    def create_help_messages(self):
        self.help_texts = {
            'presets': "Select a preconfigured scale profile.\nRight-click to manage presets.",
            'com_port': "1. Connect scale via USB/RS232\n2. Click refresh (↻) to scan ports\n3. Select detected port",
            'baud_rate': "Communication speed (bits per second)\nDefault: 9600 for Precisa",
            'data_bits': "Number of data bits per character\nPrecisa: 7, Most devices: 8",
            'parity': "Error checking method\nPrecisa: ODD, Common: NONE",
            'flow_control': "Data flow management\nXON/XOFF: Software\nHARDWARE: RTS/CTS",
            'export': "Export data to CSV\nAppends to existing file with timestamp",
            'connection': "Connect/Disconnect from scale\nVerify parameters first",
            'new_preset': "Create new preset from current settings",
            'email': "Send data via email\nRequires SMTP server configuration",
            'smtp_settings': "Configure email server settings\nFor Gmail, use an App Password"
        }

    def on_window_resize(self, event=None):
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = BalanceLogger(root)
    root.mainloop()
