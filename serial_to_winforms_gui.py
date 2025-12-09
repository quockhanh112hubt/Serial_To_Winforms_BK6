import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
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
from ftplib import FTP
import ctypes
from ctypes import wintypes

# Mutex for single instance
MUTEX_NAME = "Global\\SerialToWinFormsBK6_SingleInstance_Mutex"
mutex_handle = None

def check_single_instance():
    """Check if another instance is already running"""
    global mutex_handle
    
    # Windows API functions
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    
    # CreateMutex
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    
    # Create mutex
    mutex_handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    
    # Check if mutex already exists (ERROR_ALREADY_EXISTS = 183)
    last_error = ctypes.get_last_error()
    
    if last_error == 183:  # ERROR_ALREADY_EXISTS
        # Another instance is running
        messagebox.showerror(
            "Application Already Running",
            "Serial To WinForms is already running!\n\n"
            "Only one instance of this application can run at a time.\n"
            "Please check the system tray or taskbar for the running instance.",
            icon='error'
        )
        return False
    
    return True

def release_mutex():
    """Release the mutex on exit"""
    global mutex_handle
    if mutex_handle:
        kernel32 = ctypes.WinDLL('kernel32')
        kernel32.CloseHandle(mutex_handle)
        mutex_handle = None

# Default settings - will be loaded from settings.json
class AppSettings:
    def __init__(self):
        # Update Settings
        self.program_directory = "C:\\Serial_to_MES"
        self.ftp_server = "10.62.102.5"
        self.ftp_user = "update"
        self.ftp_password = "update"
        self.ftp_directory = "KhanhDQ/Update_Program/Serial_to_MES/"
        
        # Monitoring Settings
        self.max_log_lines = 50
        self.idle_timeout_minutes = 30
        self.max_consecutive_errors = 10
        self.connection_grace_period = 5
        self.max_disconnect_tolerance = 20
        self.auto_reset = False  # Auto reset before sending data to Shop-Flow
        
    def to_dict(self):
        return {
            'program_directory': self.program_directory,
            'ftp_server': self.ftp_server,
            'ftp_user': self.ftp_user,
            'ftp_password': self.ftp_password,
            'ftp_directory': self.ftp_directory,
            'max_log_lines': self.max_log_lines,
            'idle_timeout_minutes': self.idle_timeout_minutes,
            'max_consecutive_errors': self.max_consecutive_errors,
            'connection_grace_period': self.connection_grace_period,
            'max_disconnect_tolerance': self.max_disconnect_tolerance,
            'auto_reset': self.auto_reset
        }
    
    def from_dict(self, data):
        self.program_directory = data.get('program_directory', self.program_directory)
        self.ftp_server = data.get('ftp_server', self.ftp_server)
        self.ftp_user = data.get('ftp_user', self.ftp_user)
        self.ftp_password = data.get('ftp_password', self.ftp_password)
        self.ftp_directory = data.get('ftp_directory', self.ftp_directory)
        self.max_log_lines = data.get('max_log_lines', self.max_log_lines)
        self.idle_timeout_minutes = data.get('idle_timeout_minutes', self.idle_timeout_minutes)
        self.max_consecutive_errors = data.get('max_consecutive_errors', self.max_consecutive_errors)
        self.connection_grace_period = data.get('connection_grace_period', self.connection_grace_period)
        self.max_disconnect_tolerance = data.get('max_disconnect_tolerance', self.max_disconnect_tolerance)
        self.auto_reset = data.get('auto_reset', self.auto_reset)

# Global settings instance
app_settings = AppSettings()

# Helper functions using global settings
def get_program_directory():
    return app_settings.program_directory

def get_update_script_executable():
    return os.path.join(app_settings.program_directory, "update_script.exe")

def get_ftp_base_url():
    return f"ftp://{app_settings.ftp_user}:{app_settings.ftp_password}@{app_settings.ftp_server}/{app_settings.ftp_directory}"

def get_version_url():
    return get_ftp_base_url() + "version.txt"

def get_current_version_file():
    return os.path.join(app_settings.program_directory, "version.txt")

def get_current_version():
    current_version_file = get_current_version_file()
    if os.path.exists(current_version_file):
        with open(current_version_file, "r") as file:
            return file.read().strip()
    return "0.0.0"

def get_latest_version():
    try:
        version_url = get_version_url()
        with urllib.request.urlopen(version_url) as response:
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
    update_script = get_update_script_executable()
    process = subprocess.Popen([update_script])
    print(f"Đã khởi chạy {update_script}, PID: {process.pid}")
    sys.exit()


class SerialToWinFormsGUI:
    def __init__(self, root):
        self.root = root
        version = get_current_version()
        self.root.title("Serial To WinForms - Control Panel v"+version)
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Load settings first
        self.load_settings()
        
        # Serial handler
        self.serial_handler = None
        self.running = False
        
        # System tray
        self.tray_icon = None
        self.is_hidden = False
        
        # Setup menu bar
        self.setup_menu()
        
        # Setup GUI
        self.setup_ui()
        
        # Load config
        self.load_config()
        
        # Setup system tray
        self.setup_tray_icon()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
    
    def get_app_directory(self):
        """Get application directory (works for both .py and .exe)"""
        if getattr(sys, 'frozen', False):
            # Running as compiled exe
            return os.path.dirname(sys.executable)
        else:
            # Running as script
            return os.path.dirname(os.path.abspath(__file__))
    
    def load_settings(self):
        """Load application settings from settings.json"""
        try:
            app_dir = self.get_app_directory()
            settings_path = os.path.join(app_dir, 'settings.json')
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    app_settings.from_dict(data)
        except Exception as e:
            print(f"Failed to load settings: {e}, using defaults")
    
    def save_settings(self):
        """Save application settings to settings.json"""
        try:
            app_dir = self.get_app_directory()
            settings_path = os.path.join(app_dir, 'settings.json')
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(app_settings.to_dict(), f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")
            return False
    
    def setup_menu(self):
        """Setup menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Settings", command=self.open_settings_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit_app)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        
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
        settings_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
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
        
        # Apply settings to instance variables
        self.max_log_lines = app_settings.max_log_lines
        self.idle_timeout_minutes = app_settings.idle_timeout_minutes
        self.max_consecutive_errors = app_settings.max_consecutive_errors
        self.connection_grace_period = app_settings.connection_grace_period
        
        # Connection loss tracking
        self.serial_disconnect_count = 0
        self.winforms_disconnect_count = 0
        self.max_disconnect_tolerance = app_settings.max_disconnect_tolerance
        
    def load_config(self):
        """Load configuration from config.json"""
        try:
            app_dir = self.get_app_directory()
            config_path = os.path.join(app_dir, 'config.json')
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
            app_dir = self.get_app_directory()
            config_path = os.path.join(app_dir, 'config.json')
            
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
            
            # Create handler instance with auto_reset setting
            self.serial_handler = SerialToWinForms(auto_reset=app_settings.auto_reset)
            
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
    
    def show_about(self):
        """Show about dialog"""
        AboutDialog(self.root)
    
    def open_settings_dialog(self):
        """Open settings dialog"""
        dialog = SettingsDialog(self.root, self)
        self.root.wait_window(dialog.top)


class AboutDialog:
    """Beautiful About Dialog"""
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.title("About")
        self.top.geometry("450x450")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        
        # Center the dialog
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.top.winfo_screenheight() // 2) - (350 // 2)
        self.top.geometry(f"450x450+{x}+{y}")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame with gradient-like background
        main_frame = tk.Frame(self.top, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header frame (dark blue)
        header_frame = tk.Frame(main_frame, bg="#1e3a8a", height=120)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # App icon (using text as icon)
        icon_label = tk.Label(header_frame, text="⚡", font=('Arial', 48), 
                             bg="#1e3a8a", fg="white")
        icon_label.pack(pady=10)
        
        # App name
        name_label = tk.Label(header_frame, text="Serial To WinForms", 
                             font=('Arial', 16, 'bold'), bg="#1e3a8a", fg="white")
        name_label.pack()
        
        # Content frame
        content_frame = tk.Frame(main_frame, bg="white", padx=30, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Version
        version = get_current_version()
        version_label = tk.Label(content_frame, text=f"Version {version}", 
                                font=('Arial', 11, 'bold'), bg="white", fg="#1e3a8a")
        version_label.pack(pady=(0, 15))
        
        # Description
        desc_label = tk.Label(content_frame, 
                             text="Connects serial port data to\nWinForms application automatically", 
                             font=('Arial', 10), bg="white", fg="#666666", justify=tk.CENTER)
        desc_label.pack(pady=(0, 20))
        
        # Separator
        separator = ttk.Separator(content_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=10)
        
        # Developer info
        dev_frame = tk.Frame(content_frame, bg="white")
        dev_frame.pack(pady=10)
        
        tk.Label(dev_frame, text="👨‍💻 Developed by:", font=('Arial', 9), 
                bg="white", fg="#666666").pack()
        tk.Label(dev_frame, text="KhanhIT - IT Team", font=('Arial', 10, 'bold'), 
                bg="white", fg="#1e3a8a").pack()
        
        # Copyright
        copyright_label = tk.Label(content_frame, 
                                   text="ITM Semiconductor © 2025\nAll rights reserved", 
                                   font=('Arial', 8), bg="white", fg="#999999", justify=tk.CENTER)
        copyright_label.pack(side=tk.BOTTOM, pady=(15, 0))
        
        # Close button
        btn_frame = tk.Frame(content_frame, bg="white")
        btn_frame.pack(side=tk.BOTTOM, pady=(10, 0))
        
        close_btn = tk.Button(btn_frame, text="Close", command=self.top.destroy,
                             font=('Arial', 10), bg="#1e3a8a", fg="white",
                             padx=30, pady=8, relief=tk.FLAT, cursor="hand2",
                             activebackground="#2563eb", activeforeground="white")
        close_btn.pack()
        
        # Hover effect
        def on_enter(e):
            close_btn['bg'] = '#2563eb'
        
        def on_leave(e):
            close_btn['bg'] = '#1e3a8a'
        
        close_btn.bind("<Enter>", on_enter)
        close_btn.bind("<Leave>", on_leave)


class SettingsDialog:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.top = tk.Toplevel(parent)
        self.top.title("⚙️ Settings")
        self.top.geometry("550x700")
        self.top.resizable(False, False)
        
        # Make dialog modal
        self.top.transient(parent)
        self.top.grab_set()
        
        # Center the dialog
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.top.winfo_screenheight() // 2) - (580 // 2)
        self.top.geometry(f"550x700+{x}+{y}")
        
        # Create temporary settings copy
        self.temp_settings = AppSettings()
        self.temp_settings.from_dict(app_settings.to_dict())
        
        self.setup_ui()
        
    def setup_ui(self):
        # Header frame
        header_frame = tk.Frame(self.top, bg="#1e3a8a", height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="⚙️ Application Settings", 
                              font=('Arial', 14, 'bold'), bg="#1e3a8a", fg="white")
        title_label.pack(pady=18)
        
        # Main content frame
        content_frame = tk.Frame(self.top, bg="white")
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook (tabs) with custom style
        style = ttk.Style()
        style.theme_use('default')
        style.configure('Custom.TNotebook', background='white', borderwidth=0)
        style.configure('Custom.TNotebook.Tab', padding=[20, 10], font=('Arial', 10))
        style.map('Custom.TNotebook.Tab', background=[('selected', '#1e3a8a')], 
                 foreground=[('selected', 'white')])
        
        notebook = ttk.Notebook(content_frame, style='Custom.TNotebook')
        notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Tab 1: Update Settings
        update_frame = tk.Frame(notebook, bg="white", padx=20, pady=15)
        notebook.add(update_frame, text="  📦 Update Settings  ")
        self.setup_update_tab(update_frame)
        
        # Tab 2: Monitoring Settings
        monitoring_frame = tk.Frame(notebook, bg="white", padx=20, pady=15)
        notebook.add(monitoring_frame, text="  📊 Monitoring Settings  ")
        self.setup_monitoring_tab(monitoring_frame)
        
        # Buttons frame at bottom
        button_frame = tk.Frame(self.top, bg="#f0f0f0", pady=15)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        btn_container = tk.Frame(button_frame, bg="#f0f0f0")
        btn_container.pack()
        
        # Reset button (left)
        reset_btn = tk.Button(btn_container, text="🔄 Reset to Default", 
                             command=self.reset_to_default,
                             font=('Arial', 10), bg="#dc2626", fg="white",
                             padx=15, pady=8, relief=tk.FLAT, cursor="hand2",
                             activebackground="#b91c1c", activeforeground="white")
        reset_btn.pack(side=tk.LEFT, padx=5)
        
        # Cancel button
        cancel_btn = tk.Button(btn_container, text="Cancel", 
                              command=self.top.destroy,
                              font=('Arial', 10), bg="#6b7280", fg="white",
                              padx=25, pady=8, relief=tk.FLAT, cursor="hand2",
                              activebackground="#4b5563", activeforeground="white")
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Save button
        save_btn = tk.Button(btn_container, text="💾 Save Settings", 
                            command=self.save_settings,
                            font=('Arial', 10, 'bold'), bg="#059669", fg="white",
                            padx=25, pady=8, relief=tk.FLAT, cursor="hand2",
                            activebackground="#047857", activeforeground="white")
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Hover effects
        for btn in [reset_btn, cancel_btn, save_btn]:
            self.add_hover_effect(btn)
    
    def add_hover_effect(self, button):
        """Add hover effect to button"""
        original_bg = button['bg']
        hover_bg = button['activebackground']
        
        def on_enter(e):
            button['bg'] = hover_bg
        
        def on_leave(e):
            button['bg'] = original_bg
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
    
    def setup_update_tab(self, parent):
        """Setup update settings tab"""
        # Scrollable frame
        canvas = tk.Canvas(parent, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Program Directory Section
        section1 = tk.LabelFrame(scrollable_frame, text="  📁 Program Directory  ", 
                                font=('Arial', 10, 'bold'), bg="white", fg="#1e3a8a",
                                relief=tk.GROOVE, borderwidth=2)
        section1.pack(fill=tk.X, padx=10, pady=(10, 15))
        
        inner1 = tk.Frame(section1, bg="white", padx=15, pady=10)
        inner1.pack(fill=tk.BOTH)
        
        tk.Label(inner1, text="Installation Directory:", font=('Arial', 9), 
                bg="white", fg="#374151").pack(anchor=tk.W, pady=(0, 5))
        
        dir_frame = tk.Frame(inner1, bg="white")
        dir_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.program_dir_var = tk.StringVar(value=self.temp_settings.program_directory)
        dir_entry = tk.Entry(dir_frame, textvariable=self.program_dir_var, 
                            font=('Arial', 10), relief=tk.SOLID, borderwidth=1)
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        
        browse_btn = tk.Button(dir_frame, text="📂 Browse", command=self.browse_directory,
                              font=('Arial', 9), bg="#3b82f6", fg="white",
                              padx=12, pady=6, relief=tk.FLAT, cursor="hand2")
        browse_btn.pack(side=tk.RIGHT, padx=(5, 0))
        self.add_hover_effect(browse_btn)
        
        tk.Label(inner1, text="💡 Directory where program files are installed", 
                font=('Arial', 8), bg="white", fg="#9ca3af").pack(anchor=tk.W)
        
        # FTP Server Section
        section2 = tk.LabelFrame(scrollable_frame, text="  🌐 FTP Server Settings  ", 
                                font=('Arial', 10, 'bold'), bg="white", fg="#1e3a8a",
                                relief=tk.GROOVE, borderwidth=2)
        section2.pack(fill=tk.X, padx=10, pady=(0, 15))
        
        inner2 = tk.Frame(section2, bg="white", padx=15, pady=10)
        inner2.pack(fill=tk.BOTH)
        
        # FTP fields
        ftp_fields = [
            ("Server IP:", "ftp_server"),
            ("Username:", "ftp_user"),
            ("Password:", "ftp_password"),
            ("Directory:", "ftp_directory")
        ]
        
        for i, (label_text, var_name) in enumerate(ftp_fields):
            tk.Label(inner2, text=label_text, font=('Arial', 9), 
                    bg="white", fg="#374151").grid(row=i, column=0, sticky=tk.W, pady=8)
            
            var = tk.StringVar(value=getattr(self.temp_settings, var_name))
            setattr(self, f"{var_name}_var", var)
            
            entry = tk.Entry(inner2, textvariable=var, font=('Arial', 10),
                           relief=tk.SOLID, borderwidth=1, width=45)
            if "password" in var_name:
                entry.config(show="•")
            entry.grid(row=i, column=1, sticky=(tk.W, tk.E), pady=8, padx=(10, 0), ipady=4)
        
        inner2.columnconfigure(1, weight=1)
        
        # Test FTP button
        test_frame = tk.Frame(inner2, bg="white")
        test_frame.grid(row=len(ftp_fields), column=0, columnspan=2, pady=(10, 5))
        
        test_btn = tk.Button(test_frame, text="🔌 Test FTP Connection", 
                            command=self.test_ftp_connection,
                            font=('Arial', 9), bg="#8b5cf6", fg="white",
                            padx=15, pady=7, relief=tk.FLAT, cursor="hand2")
        test_btn.pack()
        self.add_hover_effect(test_btn)
    
    def setup_monitoring_tab(self, parent):
        """Setup monitoring settings tab"""
        # Scrollable frame
        canvas = tk.Canvas(parent, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Settings Section
        section = tk.LabelFrame(scrollable_frame, text="  ⚙️ Monitoring Parameters  ", 
                               font=('Arial', 10, 'bold'), bg="white", fg="#1e3a8a",
                               relief=tk.GROOVE, borderwidth=2)
        section.pack(fill=tk.X, padx=10, pady=10)
        
        inner = tk.Frame(section, bg="white", padx=15, pady=15)
        inner.pack(fill=tk.BOTH)
        
        settings = [
            ("📝 Max Log Lines:", "max_log_lines", 10, 1000, 
             "Maximum number of lines to keep in activity log"),
            ("⏱️ Idle Timeout (minutes):", "idle_timeout", 1, 180, 
             "Auto-stop if no data received for this duration"),
            ("❌ Max Consecutive Errors:", "max_errors", 1, 100, 
             "Auto-stop after this many consecutive errors"),
            ("⏳ Connection Grace Period (sec):", "grace_period", 1, 60, 
             "Wait time before checking for disconnections"),
            ("🔌 Max Disconnect Tolerance:", "disconnect_tolerance", 1, 100, 
             "Number of disconnect checks before auto-stop")
        ]
        
        for i, (label_text, var_name, min_val, max_val, desc) in enumerate(settings):
            # Label and value frame
            row_frame = tk.Frame(inner, bg="white")
            row_frame.grid(row=i*3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5))
            
            tk.Label(row_frame, text=label_text, font=('Arial', 9, 'bold'), 
                    bg="white", fg="#1f2937").pack(side=tk.LEFT)
            
            # Get the correct variable
            if var_name == "max_log_lines":
                var_value = self.temp_settings.max_log_lines
            elif var_name == "idle_timeout":
                var_value = self.temp_settings.idle_timeout_minutes
            elif var_name == "max_errors":
                var_value = self.temp_settings.max_consecutive_errors
            elif var_name == "grace_period":
                var_value = self.temp_settings.connection_grace_period
            elif var_name == "disconnect_tolerance":
                var_value = self.temp_settings.max_disconnect_tolerance
            
            var = tk.IntVar(value=var_value)
            setattr(self, f"{var_name}_var", var)
            
            # Spinbox with better styling
            spinbox_frame = tk.Frame(row_frame, bg="white")
            spinbox_frame.pack(side=tk.RIGHT)
            
            spinbox = tk.Spinbox(spinbox_frame, from_=min_val, to=max_val, 
                               textvariable=var, font=('Arial', 10),
                               width=10, relief=tk.SOLID, borderwidth=1,
                               buttonbackground="#3b82f6", justify=tk.CENTER)
            spinbox.pack()
            
            # Description
            tk.Label(inner, text=f"  💡 {desc}", font=('Arial', 8), 
                    bg="white", fg="#6b7280").grid(row=i*3+1, column=0, columnspan=2, 
                                                   sticky=tk.W, pady=(0, 5))
            
            # Separator
            if i < len(settings) - 1:
                sep = ttk.Separator(inner, orient='horizontal')
                sep.grid(row=i*3+2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=8)
        
        inner.columnconfigure(0, weight=1)
        
        # Auto Reset Section
        auto_reset_section = tk.LabelFrame(scrollable_frame, text="  🔄 Auto Reset Settings  ", 
                                font=('Arial', 10, 'bold'), bg="white", fg="#1e3a8a",
                                relief=tk.GROOVE, borderwidth=2)
        auto_reset_section.pack(fill=tk.X, padx=10, pady=(0, 15))
        
        auto_reset_inner = tk.Frame(auto_reset_section, bg="white", padx=15, pady=10)
        auto_reset_inner.pack(fill=tk.BOTH)
        
        # Auto Reset checkbox
        self.auto_reset_var = tk.BooleanVar(value=self.temp_settings.auto_reset)
        
        checkbox_frame = tk.Frame(auto_reset_inner, bg="white")
        checkbox_frame.pack(fill=tk.X, pady=5)
        
        auto_reset_check = tk.Checkbutton(
            checkbox_frame,
            text="Enable Auto Reset before sending data",
            variable=self.auto_reset_var,
            font=('Arial', 10, 'bold'),
            bg="white",
            fg="#1f2937",
            activebackground="white",
            activeforeground="#1e3a8a",
            selectcolor="white",
            cursor="hand2"
        )
        auto_reset_check.pack(side=tk.LEFT)
        
        # Description
        desc_frame = tk.Frame(auto_reset_inner, bg="#fef3c7", relief=tk.SOLID, borderwidth=1)
        desc_frame.pack(fill=tk.X, pady=(10, 0))
        
        desc_inner = tk.Frame(desc_frame, bg="#fef3c7", padx=12, pady=10)
        desc_inner.pack(fill=tk.BOTH)
        
        tk.Label(desc_inner, text="📌 How Auto Reset works:", 
                font=('Arial', 9, 'bold'), bg="#fef3c7", fg="#92400e").pack(anchor=tk.W, pady=(0, 5))
        
        desc_texts = [
            "✓ Enabled: Automatically click Reset button before sending each data to Shop-Flow",
            "✗ Disabled: Send data directly without resetting (default behavior)",
            "⚠️ This is independent from receiving 'RESET' command via serial port"
        ]
        
        for text in desc_texts:
            tk.Label(desc_inner, text=text, font=('Arial', 8), 
                    bg="#fef3c7", fg="#78350f", justify=tk.LEFT).pack(anchor=tk.W, pady=2)
        
        # Info box
        info_frame = tk.Frame(scrollable_frame, bg="#eff6ff", relief=tk.SOLID, borderwidth=1)
        info_frame.pack(fill=tk.X, padx=10, pady=15)
        
        info_inner = tk.Frame(info_frame, bg="#eff6ff", padx=15, pady=12)
        info_inner.pack(fill=tk.BOTH)
        
        tk.Label(info_inner, text="ℹ️ Important Information", 
                font=('Arial', 10, 'bold'), bg="#eff6ff", fg="#1e40af").pack(anchor=tk.W, pady=(0, 8))
        
        info_text = [
            "• Lower values = more aggressive auto-stop",
            "• Higher values = more tolerance for temporary issues",
            "• Changes take effect immediately for new connections",
            "• Restart required for some monitoring parameters"
        ]
        
        for text in info_text:
            tk.Label(info_inner, text=text, font=('Arial', 9), 
                    bg="#eff6ff", fg="#1e40af", justify=tk.LEFT).pack(anchor=tk.W, pady=2)
    
    def browse_directory(self):
        """Browse for program directory"""
        directory = filedialog.askdirectory(
            title="Select Program Directory",
            initialdir=self.program_dir_var.get()
        )
        if directory:
            self.program_dir_var.set(directory)
    
    def test_ftp_connection(self):
        """Test FTP connection"""
        try:
            # Show progress
            test_window = tk.Toplevel(self.top)
            test_window.title("Testing FTP Connection")
            test_window.geometry("300x100")
            test_window.transient(self.top)
            test_window.grab_set()
            
            ttk.Label(test_window, text="Testing FTP connection...", font=('Arial', 10)).pack(pady=20)
            progress = ttk.Progressbar(test_window, mode='indeterminate', length=200)
            progress.pack(pady=10)
            progress.start()
            
            # Test connection in thread
            def test():
                try:
                    ftp = FTP(self.ftp_server_var.get())
                    ftp.login(self.ftp_user_var.get(), self.ftp_password_var.get())
                    ftp.cwd(self.ftp_directory_var.get())
                    ftp.quit()
                    
                    test_window.destroy()
                    messagebox.showinfo("Success", "FTP connection successful!")
                except Exception as e:
                    test_window.destroy()
                    messagebox.showerror("Error", f"FTP connection failed:\n{e}")
            
            threading.Thread(target=test, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to test FTP connection:\n{e}")
    
    def reset_to_default(self):
        """Reset all settings to default values"""
        if messagebox.askyesno("Confirm Reset", "Reset all settings to default values?"):
            # Reset to defaults
            default_settings = AppSettings()
            
            # Update UI
            self.program_dir_var.set(default_settings.program_directory)
            self.ftp_server_var.set(default_settings.ftp_server)
            self.ftp_user_var.set(default_settings.ftp_user)
            self.ftp_password_var.set(default_settings.ftp_password)
            self.ftp_directory_var.set(default_settings.ftp_directory)
            self.max_log_lines_var.set(default_settings.max_log_lines)
            self.idle_timeout_var.set(default_settings.idle_timeout_minutes)
            self.max_errors_var.set(default_settings.max_consecutive_errors)
            self.grace_period_var.set(default_settings.connection_grace_period)
            self.disconnect_tolerance_var.set(default_settings.max_disconnect_tolerance)
            
            messagebox.showinfo("Reset Complete", "All settings have been reset to default values.")
    
    def validate_settings(self):
        """Validate settings before saving"""
        # Validate numeric values
        if self.max_log_lines_var.get() < 10:
            messagebox.showerror("Validation Error", "Max log lines must be at least 10")
            return False
        
        if self.idle_timeout_var.get() < 1:
            messagebox.showerror("Validation Error", "Idle timeout must be at least 1 minute")
            return False
        
        if self.max_errors_var.get() < 1:
            messagebox.showerror("Validation Error", "Max consecutive errors must be at least 1")
            return False
        
        if self.grace_period_var.get() < 1:
            messagebox.showerror("Validation Error", "Connection grace period must be at least 1 second")
            return False
        
        if self.disconnect_tolerance_var.get() < 1:
            messagebox.showerror("Validation Error", "Max disconnect tolerance must be at least 1")
            return False
        
        # Validate FTP settings
        if not self.ftp_server_var.get().strip():
            messagebox.showerror("Validation Error", "FTP server cannot be empty")
            return False
        
        if not self.ftp_user_var.get().strip():
            messagebox.showerror("Validation Error", "FTP username cannot be empty")
            return False
        
        return True
    
    def save_settings(self):
        """Save settings and close dialog"""
        if not self.validate_settings():
            return
        
        # Update global settings
        app_settings.program_directory = self.program_dir_var.get()
        app_settings.ftp_server = self.ftp_server_var.get()
        app_settings.ftp_user = self.ftp_user_var.get()
        app_settings.ftp_password = self.ftp_password_var.get()
        app_settings.ftp_directory = self.ftp_directory_var.get()
        app_settings.max_log_lines = self.max_log_lines_var.get()
        app_settings.idle_timeout_minutes = self.idle_timeout_var.get()
        app_settings.max_consecutive_errors = self.max_errors_var.get()
        app_settings.connection_grace_period = self.grace_period_var.get()
        app_settings.max_disconnect_tolerance = self.disconnect_tolerance_var.get()
        app_settings.auto_reset = self.auto_reset_var.get()
        
        # Save to file
        if self.app.save_settings():
            # Update app instance variables
            self.app.max_log_lines = app_settings.max_log_lines
            self.app.idle_timeout_minutes = app_settings.idle_timeout_minutes
            self.app.max_consecutive_errors = app_settings.max_consecutive_errors
            self.app.connection_grace_period = app_settings.connection_grace_period
            self.app.max_disconnect_tolerance = app_settings.max_disconnect_tolerance
            
            # Update auto_reset in serial handler if it exists
            if hasattr(self.app, 'serial_handler') and self.app.serial_handler:
                self.app.serial_handler.auto_reset = app_settings.auto_reset
                self.app.log_message(f"Auto Reset updated: {'Enabled' if app_settings.auto_reset else 'Disabled'}", "INFO")
            
            messagebox.showinfo("Success", "Settings saved successfully!\n\nNote: Some settings may require restarting the application to take full effect.")
            self.top.destroy()
        else:
            messagebox.showerror("Error", "Failed to save settings. Please try again.")

def main():
    # Check for single instance before creating GUI
    if not check_single_instance():
        sys.exit(1)
    
    root = tk.Tk()
    app = SerialToWinFormsGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Release mutex on exit
    def on_app_exit():
        release_mutex()
        root.quit()
    
    root.protocol("WM_DESTROY", on_app_exit)
    root.mainloop()

if __name__ == "__main__":
    try:
        # check_for_updates()
        main()
    finally:
        release_mutex()
