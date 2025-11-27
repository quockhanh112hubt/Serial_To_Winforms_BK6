import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import logging
from datetime import datetime, timedelta
import os
import sys
import json
from serial_to_winforms_bk6 import SerialToWinForms
import pystray
from PIL import Image, ImageDraw
import time
import urllib.request
import subprocess

PROGRAM_DIRECTORY = "C:\\Serial_to_MES"
UPDATE_SCRIPT_EXECUTABLE = os.path.join(PROGRAM_DIRECTORY, "update_script.exe")

#FTP Server
FTP_BASE_URL = "ftp://update:update@10.62.102.5/KhanhDQ/Update_Program/Serial_to_MES/"
VERSION_URL = FTP_BASE_URL + "version.txt"


CURRENT_VERSION_FILE = "C:\\Serial_to_MES\\version.txt"

def get_current_version():
    if os.path.exists(CURRENT_VERSION_FILE):
        with open(CURRENT_VERSION_FILE, "r") as file:
            return file.read().strip()
    return "0.0.0"

def get_latest_version():
    try:
        with urllib.request.urlopen(VERSION_URL) as response:
            latest_version = response.read().decode('utf-8').strip()
        return latest_version
    except Exception as e:
        print(f"Không thể lấy phiên bản mới nhất: {e}")
        return None

def check_for_updates():
    current_version = get_current_version()
    latest_version = get_latest_version()
    
    if latest_version and latest_version > current_version:
        initiate_update()

def initiate_update():
    print("Đang chuẩn bị cập nhật và khởi động lại chương trình...")
    process = subprocess.Popen([UPDATE_SCRIPT_EXECUTABLE])
    print(f"Đã khởi chạy {UPDATE_SCRIPT_EXECUTABLE}, PID: {process.pid}")
    sys.exit()


class SerialToWinFormsGUI:
    def __init__(self, root):
        self.root = root
        version = get_current_version()
        self.root.title("Serial To WinForms - Control Panel v"+version)
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Serial handler
        self.serial_handler = None
        self.running = False
        
        # System tray
        self.tray_icon = None
        self.is_hidden = False
        
        # Setup GUI
        self.setup_ui()
        
        # Load config
        self.load_config()
        
        # Setup system tray
        self.setup_tray_icon()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        version = get_current_version()
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Serial To WinForms Control Panel v"+version, 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Settings Frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # COM Port
        ttk.Label(settings_frame, text="COM Port:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.port_var = tk.StringVar(value="COM10")
        self.port_entry = ttk.Entry(settings_frame, textvariable=self.port_var, width=15)
        self.port_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Baudrate
        ttk.Label(settings_frame, text="Baudrate:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.baudrate_var = tk.StringVar(value="9600")
        self.baudrate_combo = ttk.Combobox(settings_frame, textvariable=self.baudrate_var, 
                                          values=["9600", "19200", "38400", "57600", "115200"], 
                                          width=12, state="readonly")
        self.baudrate_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Target App
        ttk.Label(settings_frame, text="Target App:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_app_var = tk.StringVar(value="Shop-Flow System From Vietnam(Pack)")
        self.target_app_entry = ttk.Entry(settings_frame, textvariable=self.target_app_var, width=40)
        self.target_app_entry.grid(row=1, column=1, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Textbox ID
        ttk.Label(settings_frame, text="Textbox ID:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.textbox_id_var = tk.StringVar(value="GIFTBOX_AUTO")
        self.textbox_id_entry = ttk.Entry(settings_frame, textvariable=self.textbox_id_var, width=20)
        self.textbox_id_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Status Frame
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Serial Status
        ttk.Label(status_frame, text="Serial Port:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.serial_status_label = ttk.Label(status_frame, text="●", font=('Arial', 20), foreground="red")
        self.serial_status_label.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.serial_status_text = ttk.Label(status_frame, text="Disconnected", foreground="red")
        self.serial_status_text.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        # WinForms Status
        ttk.Label(status_frame, text="Shop-Flow:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.winforms_status_label = ttk.Label(status_frame, text="●", font=('Arial', 20), foreground="red")
        self.winforms_status_label.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.winforms_status_text = ttk.Label(status_frame, text="Disconnected", foreground="red")
        self.winforms_status_text.grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        
        # Last Data Received
        ttk.Label(status_frame, text="Last Data:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.last_data_label = ttk.Label(status_frame, text="N/A", foreground="gray")
        self.last_data_label.grid(row=2, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Control Buttons Frame
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Start/Stop Button
        self.start_stop_btn = ttk.Button(control_frame, text="Start", 
                                         command=self.toggle_start_stop, width=15)
        self.start_stop_btn.grid(row=0, column=0, padx=5)
        
        # Save Config Button
        self.save_config_btn = ttk.Button(control_frame, text="Save Config", 
                                          command=self.save_config, width=15)
        self.save_config_btn.grid(row=0, column=1, padx=5)
        
        # Clear Log Button
        self.clear_log_btn = ttk.Button(control_frame, text="Clear Log", 
                                        command=self.clear_log, width=15)
        self.clear_log_btn.grid(row=0, column=2, padx=5)
        
        # Data Counter Frame
        counter_frame = ttk.Frame(main_frame)
        counter_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
        ttk.Label(counter_frame, text="Data Received:").grid(row=0, column=0, padx=5)
        self.data_counter_var = tk.StringVar(value="0")
        self.data_counter_label = ttk.Label(counter_frame, textvariable=self.data_counter_var, 
                                            font=('Arial', 12, 'bold'), foreground="blue")
        self.data_counter_label.grid(row=0, column=1, padx=5)
        
        ttk.Label(counter_frame, text="Success:").grid(row=0, column=2, padx=5)
        self.success_counter_var = tk.StringVar(value="0")
        self.success_counter_label = ttk.Label(counter_frame, textvariable=self.success_counter_var, 
                                               font=('Arial', 12, 'bold'), foreground="green")
        self.success_counter_label.grid(row=0, column=3, padx=5)
        
        ttk.Label(counter_frame, text="Errors:").grid(row=0, column=4, padx=5)
        self.error_counter_var = tk.StringVar(value="0")
        self.error_counter_label = ttk.Label(counter_frame, textvariable=self.error_counter_var, 
                                             font=('Arial', 12, 'bold'), foreground="red")
        self.error_counter_label.grid(row=0, column=5, padx=5)
        
        # Log Frame
        log_frame = ttk.LabelFrame(main_frame, text="Activity Log", padding="5")
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log Text Area
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure text tags for colored output
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("WARNING", foreground="orange")
        
        # Counters
        self.data_count = 0
        self.success_count = 0
        self.error_count = 0
        
        # Auto-stop tracking
        self.last_data_time = None
        self.consecutive_errors = 0
        self.max_log_lines = 50  # Maximum lines in activity log
        self.idle_timeout_minutes = 30  # Auto-stop if no data for 30 minutes
        self.max_consecutive_errors = 10  # Auto-stop if more than 10 consecutive errors
        self.connection_grace_period = 5  # Don't auto-stop in first 5 seconds after start
        
        # Connection loss tracking
        self.serial_disconnect_count = 0
        self.winforms_disconnect_count = 0
        self.max_disconnect_tolerance = 20  # Allow 20 consecutive disconnect detections before stopping
        
    def load_config(self):
        """Load configuration from config.json"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), 'config.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                self.port_var.set(config.get('port', 'COM10'))
                self.baudrate_var.set(str(config.get('baudrate', 9600)))
                self.target_app_var.set(config.get('target_app_title', 'Shop-Flow System From Vietnam(Pack)'))
                self.textbox_id_var.set(config.get('textbox_auto_id', 'GIFTBOX_AUTO'))
                self.log_message("Config loaded successfully", "SUCCESS")
            else:
                self.log_message("Config file not found, using defaults", "WARNING")
        except Exception as e:
            self.log_message(f"Failed to load config: {e}", "ERROR")
            messagebox.showerror("Config Error", f"Failed to load config:\n{e}\n\nUsing default values.")
            
    def save_config(self):
        """Save configuration to config.json"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), 'config.json')
            
            # Ensure directory exists
            config_dir = os.path.dirname(config_path)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir, exist_ok=True)
            
            config = {
                'port': self.port_var.get(),
                'baudrate': int(self.baudrate_var.get()),
                'target_app_title': self.target_app_var.get(),
                'textbox_auto_id': self.textbox_id_var.get(),
                'backend': 'win32'
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            self.log_message("Config saved successfully", "SUCCESS")
            # messagebox.showinfo("Success", "Configuration saved successfully!")
        except Exception as e:
            self.log_message(f"Failed to save config: {e}", "ERROR")
            messagebox.showerror("Error", f"Failed to save config: {e}")
    
    def toggle_start_stop(self):
        """Start or stop the serial handler"""
        if not self.running:
            self.start_handler()
        else:
            self.stop_handler()
    
    def start_handler(self):
        """Start the serial to winforms handler"""
        try:
            # Validate settings first
            if not self.port_var.get():
                messagebox.showerror("Error", "Please enter COM port")
                return
            
            # Disable settings while running
            self.port_entry.config(state='disabled')
            self.baudrate_combo.config(state='disabled')
            self.target_app_entry.config(state='disabled')
            self.textbox_id_entry.config(state='disabled')
            
            # Save config before starting (but don't fail if it errors)
            try:
                self.save_config()
            except Exception as save_err:
                self.log_message(f"Warning: Could not save config: {save_err}", "WARNING")
            
            # Create handler instance
            self.serial_handler = SerialToWinForms()
            
            # Setup custom logging to GUI
            self.setup_gui_logging()
            
            # Start in background thread
            self.running = True
            self.thread = threading.Thread(target=self.run_handler, daemon=True)
            self.thread.start()
            
            # Reset auto-stop tracking
            self.last_data_time = datetime.now()
            self.consecutive_errors = 0
            self.start_time = datetime.now()  # Track when handler started
            self.serial_disconnect_count = 0  # Reset disconnect counters
            self.winforms_disconnect_count = 0
            
            # Update UI
            self.start_stop_btn.config(text="Stop")
            self.log_message("Starting Serial To WinForms handler...", "INFO")
            
            # Start status monitoring
            self.monitor_status()
            
        except Exception as e:
            self.log_message(f"Failed to start: {e}", "ERROR")
            messagebox.showerror("Error", f"Failed to start: {e}")
            self.stop_handler()
    
    def stop_handler(self):
        """Stop the serial handler"""
        try:
            self.running = False
            if self.serial_handler:
                self.serial_handler.running = False
                self.serial_handler = None
            
            # Update UI
            self.start_stop_btn.config(text="Start")
            self.update_status("serial", False)
            self.update_status("winforms", False)
            
            # Update tray icon to gray (stopped state)
            self.update_tray_icon()
            
            # Enable settings
            self.port_entry.config(state='normal')
            self.baudrate_combo.config(state='normal')
            self.target_app_entry.config(state='normal')
            self.textbox_id_entry.config(state='normal')
            
            self.log_message("Handler stopped", "WARNING")
            
        except Exception as e:
            self.log_message(f"Error stopping handler: {e}", "ERROR")
    
    def run_handler(self):
        """Run the serial handler in background thread"""
        try:
            self.serial_handler.start()
            
            # Wait longer for connections to establish properly
            time.sleep(3)
            
            # Check if both connections are successful
            serial_ok = self.serial_handler.serial_conn is not None and self.serial_handler.serial_conn.is_open
            winforms_ok = self.serial_handler.window is not None and self.serial_handler.textbox is not None
            
            if not serial_ok:
                self.log_message("❌ Failed to connect to Serial Port", "ERROR")
                self.root.after(0, lambda: messagebox.showerror("Connection Failed", 
                    f"Cannot connect to Serial Port: {self.port_var.get()}\n\nPlease check:\n• COM port exists and is available\n• Port not in use by another program\n• Baudrate is correct ({self.baudrate_var.get()})"))
                self.root.after(0, self.stop_handler)
                return
            
            if not winforms_ok:
                self.log_message("❌ Failed to connect to Shop-Flow", "ERROR")
                self.root.after(0, lambda: messagebox.showerror("Connection Failed", 
                    f"Cannot connect to Shop-Flow!\n\nPlease check:\n• Shop-Flow application is running\n• Target App title: {self.target_app_var.get()}\n• Textbox ID: {self.textbox_id_var.get()}"))
                self.root.after(0, self.stop_handler)
                return
            
            # Both connections OK
            self.log_message("✅ All connections established successfully", "SUCCESS")
            
        except Exception as e:
            self.log_message(f"Handler error: {e}", "ERROR")
            self.root.after(0, self.stop_handler)
    
    def setup_gui_logging(self):
        """Setup logging to output to GUI"""
        class GUIHandler(logging.Handler):
            def __init__(self, gui):
                super().__init__()
                self.gui = gui
                
            def emit(self, record):
                msg = self.format(record)
                level = record.levelname
                
                # Determine tag based on level
                if level == "ERROR":
                    tag = "ERROR"
                    self.gui.error_count += 1
                    self.gui.consecutive_errors += 1  # Track consecutive errors
                    self.gui.root.after(0, lambda: self.gui.error_counter_var.set(str(self.gui.error_count)))
                    
                    # Check for too many consecutive errors
                    if self.gui.consecutive_errors >= self.gui.max_consecutive_errors:
                        self.gui.root.after(0, lambda: self.gui.auto_stop_due_to_errors())
                        
                elif level == "WARNING":
                    tag = "WARNING"
                elif "successful" in msg.lower() or "connected" in msg.lower():
                    tag = "SUCCESS"
                    self.gui.success_count += 1
                    self.gui.consecutive_errors = 0  # Reset on success
                    self.gui.root.after(0, lambda: self.gui.success_counter_var.set(str(self.gui.success_count)))
                else:
                    tag = "INFO"
                    self.gui.consecutive_errors = 0  # Reset on normal info
                
                # Check for data received
                if "Raw serial data received:" in msg:
                    self.gui.data_count += 1
                    self.gui.last_data_time = datetime.now()  # Update last data time
                    self.gui.consecutive_errors = 0  # Reset on data received
                    self.gui.root.after(0, lambda: self.gui.data_counter_var.set(str(self.gui.data_count)))
                    # Extract and display last data
                    try:
                        data = msg.split("Raw serial data received:")[1].split("(")[0].strip()
                        self.gui.root.after(0, lambda d=data: self.gui.last_data_label.config(text=d[:50]))
                    except:
                        pass
                
                self.gui.root.after(0, lambda: self.gui.log_message(msg, tag))
        
        # Add GUI handler to root logger
        gui_handler = GUIHandler(self)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                                   datefmt='%H:%M:%S'))
        logging.getLogger().addHandler(gui_handler)
    
    def monitor_status(self):
        """Monitor connection status"""
        if self.running and self.serial_handler:
            # Debug log every 10 seconds
            time_since_start = (datetime.now() - self.start_time).total_seconds() if hasattr(self, 'start_time') else 999
            # if int(time_since_start) % 10 == 0:
                # self.log_message(f"Monitor check at {int(time_since_start)}s", "INFO")
            
            # Check serial connection
            serial_connected = self.serial_handler.serial_conn is not None and self.serial_handler.serial_conn.is_open
            self.update_status("serial", serial_connected)
            
            # Check winforms connection - use lightweight check
            winforms_connected = False
            if self.serial_handler.window is not None and self.serial_handler.textbox is not None:
                try:
                    # Only check if window exists (lightweight, doesn't wait for response)
                    winforms_connected = self.serial_handler.window.exists(timeout=0.1)
                except Exception as e:
                    # Any exception means window is not accessible (likely closed)
                    winforms_connected = False
                    
            self.update_status("winforms", winforms_connected)
            
            # Only check for disconnection after grace period (to allow initial connection time)
            time_since_start = (datetime.now() - self.start_time).total_seconds() if hasattr(self, 'start_time') else 999
            
            if time_since_start > self.connection_grace_period:
                # Check serial connection with tolerance
                if not serial_connected:
                    self.serial_disconnect_count += 1
                    if self.serial_disconnect_count >= self.max_disconnect_tolerance:
                        self.log_message(f"⚠️ Serial Port disconnected ({self.serial_disconnect_count} times) - stopping handler", "ERROR")
                        messagebox.showerror("Connection Lost", f"Serial Port disconnected {self.serial_disconnect_count} times!\nHandler will be stopped.")
                        self.stop_handler()
                        return
                    else:
                        self.log_message(f"⚠️ Serial Port disconnect detected ({self.serial_disconnect_count}/{self.max_disconnect_tolerance})", "WARNING")
                else:
                    # Reset counter if connection is back
                    if self.serial_disconnect_count > 0:
                        self.log_message(f"✅ Serial Port reconnected (reset counter)", "SUCCESS")
                    self.serial_disconnect_count = 0
                
                # Check Shop-Flow connection with tolerance
                if not winforms_connected:
                    self.winforms_disconnect_count += 1
                    if self.winforms_disconnect_count >= self.max_disconnect_tolerance:
                        self.log_message(f"⚠️ Shop-Flow disconnected ({self.winforms_disconnect_count} times) - stopping handler", "ERROR")
                        messagebox.showerror("Connection Lost", f"Shop-Flow connection lost {self.winforms_disconnect_count} times!\nHandler will be stopped.")
                        self.stop_handler()
                        return
                    else:
                        self.log_message(f"⚠️ Shop-Flow disconnect detected ({self.winforms_disconnect_count}/{self.max_disconnect_tolerance})", "WARNING")
                else:
                    # Reset counter if connection is back
                    if self.winforms_disconnect_count > 0:
                        self.log_message(f"✅ Shop-Flow reconnected (reset counter)", "SUCCESS")
                    self.winforms_disconnect_count = 0
            
            # Update tray icon
            self.update_tray_icon()
            
            # Check for idle timeout (no data received for too long)
            if self.last_data_time and time_since_start > self.connection_grace_period:
                idle_duration = datetime.now() - self.last_data_time
                if idle_duration > timedelta(minutes=self.idle_timeout_minutes):
                    self.auto_stop_due_to_idle()
                    return  # Don't schedule next check
            
            # Schedule next check
            self.root.after(1000, self.monitor_status)
    
    def update_status(self, status_type, connected):
        """Update status indicators"""
        if status_type == "serial":
            if connected:
                self.serial_status_label.config(foreground="green")
                self.serial_status_text.config(text="Connected", foreground="green")
            else:
                self.serial_status_label.config(foreground="red")
                self.serial_status_text.config(text="Disconnected", foreground="red")
        elif status_type == "winforms":
            if connected:
                self.winforms_status_label.config(foreground="green")
                self.winforms_status_text.config(text="Connected", foreground="green")
            else:
                self.winforms_status_label.config(foreground="red")
                self.winforms_status_text.config(text="Disconnected", foreground="red")
    
    def log_message(self, message, tag="INFO"):
        """Add message to log text area"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        self.log_text.see(tk.END)  # Auto-scroll to bottom
        
        # Limit log lines to max_log_lines
        lines = int(self.log_text.index('end-1c').split('.')[0])
        if lines > self.max_log_lines:
            # Delete oldest lines
            excess_lines = lines - self.max_log_lines
            self.log_text.delete('1.0', f'{excess_lines}.0')
    
    def clear_log(self):
        """Clear the log text area"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Log cleared", "INFO")
    
    def auto_stop_due_to_idle(self):
        """Auto-stop handler due to no data timeout"""
        self.log_message(f"⚠️ AUTO-STOP: No data received for {self.idle_timeout_minutes} minutes", "WARNING")
        messagebox.showwarning("Auto-Stop", 
                              f"Handler stopped automatically:\nNo data received for {self.idle_timeout_minutes} minutes")
        self.stop_handler()
    
    def auto_stop_due_to_errors(self):
        """Auto-stop handler due to too many consecutive errors"""
        self.log_message(f"⚠️ AUTO-STOP: Too many consecutive errors ({self.consecutive_errors})", "ERROR")
        messagebox.showerror("Auto-Stop", 
                            f"Handler stopped automatically:\nToo many consecutive errors ({self.consecutive_errors})")
        self.stop_handler()
    
    def create_tray_image(self):
        """Create icon for system tray"""
        # Create a simple icon image
        width = 64
        height = 64
        color1 = "green" if self.running else "gray"
        
        image = Image.new('RGB', (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle(
            (width // 4, height // 4, width * 3 // 4, height * 3 // 4),
            fill="white")
        
        return image
    
    def setup_tray_icon(self):
        """Setup system tray icon"""
        try:
            # Create menu
            menu = pystray.Menu(
                pystray.MenuItem("Show", self.show_window, default=True),
                pystray.MenuItem("Hide", self.hide_to_tray),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Start" if not self.running else "Stop", self.tray_toggle_handler),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Exit", self.quit_app)
            )
            
            # Create icon
            image = self.create_tray_image()
            self.tray_icon = pystray.Icon("SerialToWinForms", image, "Serial To WinForms", menu)
            
            # Run icon in separate thread
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            
        except Exception as e:
            self.log_message(f"Failed to create system tray icon: {e}", "WARNING")
    
    def update_tray_icon(self):
        """Update tray icon image based on status"""
        if self.tray_icon:
            try:
                self.tray_icon.icon = self.create_tray_image()
            except:
                pass
    
    def show_window(self, icon=None, item=None):
        """Show the main window"""
        self.is_hidden = False
        self.root.after(0, self.root.deiconify)
        self.root.after(0, self.root.lift)
        self.root.after(0, self.root.focus_force)
    
    def hide_to_tray(self):
        """Hide window to system tray"""
        self.is_hidden = True
        self.root.withdraw()
        if self.tray_icon:
            try:
                self.tray_icon.notify("Serial To WinForms", "Application minimized to tray")
            except:
                pass
    
    def tray_toggle_handler(self, icon=None, item=None):
        """Toggle start/stop from tray menu"""
        self.root.after(0, self.toggle_start_stop)
    
    def quit_app(self, icon=None, item=None):
        """Quit application completely"""
        if self.running:
            self.stop_handler()
        
        # Stop tray icon
        if self.tray_icon:
            self.tray_icon.stop()
        
        # Destroy window
        self.root.after(0, self.root.destroy)
    
    def on_closing(self):
        """Handle window closing - minimize to tray instead of quit"""
        self.hide_to_tray()

def main():
    root = tk.Tk()
    app = SerialToWinFormsGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    check_for_updates()
    main()
