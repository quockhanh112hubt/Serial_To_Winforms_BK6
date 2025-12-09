import os
import serial
import serial.tools.list_ports
import threading
import time
import pywinauto
import logging
import json
import pystray
from PIL import Image, ImageDraw

import datetime
import tkinter.messagebox as messagebox

# Logging setup
os.makedirs('log', exist_ok=True)
log_filename = f"log/{datetime.date.today()}.txt"
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[logging.FileHandler(log_filename, encoding='utf-8'), logging.StreamHandler()])

class SerialToWinForms:
    def __init__(self, auto_reset=False):
        # Load config from JSON
        try:
            with open('config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
        except FileNotFoundError:
            logging.warning("config.json not found, using defaults")
            config = {}
        
        self.port = config.get('port', 'COM7')
        self.baudrate = int(config.get('baudrate', 115200))
        self.target_app_title = config.get('target_app_title', 'Shop-Flow System From Indonesia(Pack)')
        self.textbox_auto_id = config.get('textbox_auto_id', 'GIFTBOX_AUTO')
        self.backend = config.get('backend', 'win32')
        self.serial_conn = None
        self.running = False
        self.app = None
        self.window = None
        self.textbox = None
        self.auto_reset = auto_reset  # Auto reset before sending data

    def list_available_ports(self):
        """List available serial ports"""
        ports = serial.tools.list_ports.comports()
        available_ports = [port.device for port in ports]
        logging.info(f"Available serial ports: {available_ports}")
        return available_ports

    def connect_serial(self):
        try:
            # Check available ports first
            available_ports = self.list_available_ports()
            if self.port not in available_ports:
                logging.error(f"Port {self.port} not found. Available ports: {available_ports}")
                if available_ports:
                    logging.info(f"Try using one of these ports: {', '.join(available_ports)}")
                return False
            
            self.serial_conn = serial.Serial(self.port, self.baudrate, timeout=1)
            logging.info(f"Serial port {self.port} connected successfully")
        except serial.SerialException as e:
            logging.error(f"Serial port connection failed: {e}")
            if "Access is denied" in str(e) or "액세스가 거부되었습니다" in str(e):
                messagebox.showerror("Error", f"Serial port connection failed: {e}")
            return False
        return True

    def parse_stx_etx_data(self, raw_data):
        """
        Parse dữ liệu theo format STX...ETX
        Input: "STXA01;A02;A03;...A20ETX"  
        Output: "A01;A02;A03;...A20"
        """
        try:
            # Kiểm tra có STX và ETX không
            if raw_data.startswith('STX') and raw_data.endswith('ETX'):
                # Lấy phần data ở giữa (bỏ STX và ETX)
                core_data = raw_data[3:-3]  # Bỏ 3 ký tự đầu (STX) và 3 ký tự cuối (ETX)
                logging.info(f"Parsed STX/ETX format: '{raw_data}' → '{core_data}'")
                return core_data
            else:
                # Nếu không có STX/ETX thì trả về data gốc
                logging.info(f"Raw data (no STX/ETX): '{raw_data}'")
                return raw_data
        except Exception as e:
            logging.error(f"Error parsing STX/ETX: {e}")
            return raw_data

    def read_serial_data(self):
        while self.running:
            if self.serial_conn:
                try:
                    start_time = time.time()
                    raw_data = self.serial_conn.readline().decode('utf-8').strip()
                    read_time = time.time() - start_time
                    if raw_data:
                        logging.info(f"Raw serial data received: {raw_data} (Read time: {read_time:.3f}s)")
                        
                        # Parse STX/ETX format first
                        parsed_data = self.parse_stx_etx_data(raw_data)
                        
                        # Check for RESET command (after parsing)
                        if parsed_data.upper() == "RESET":
                            logging.info("🔄 RESET command received from serial")
                            reset_start = time.time()
                            self.click_reset_button()
                            reset_time = time.time() - reset_start
                            logging.info(f"✅ Reset button clicked (Time: {reset_time:.3f}s)")
                            continue
                        
                        # Send entire string to Shop-Flow (no splitting)
                        input_start = time.time()
                        self.input_to_winforms(parsed_data)
                        input_time = time.time() - input_start
                        logging.info(f"WinForms input completed (Input time: {input_time:.3f}s)")
                    else:
                        logging.debug("No data, timeout")
                except Exception as e:
                    logging.error(f"Data read error: {e}")
            else:
                # Wait if no serial connection
                time.sleep(1)

    def input_to_winforms(self, data):
        if not self.textbox:
            logging.error("Textbox not initialized")
            return
        try:
            # Auto reset before sending data if enabled
            if self.auto_reset:
                logging.info("🔄 Auto Reset enabled - Resetting before sending data")
                reset_start = time.time()
                self.click_reset_button()
                reset_time = time.time() - reset_start
                logging.info(f"✅ Auto reset completed (Time: {reset_time:.3f}s)")
                time.sleep(0.5)  # Wait a bit after reset
            
            # Try method 1: set_text() + type_keys Enter
            try:
                logging.info(f"Attempting to input data: '{data}'")
                self.window.set_focus()  # Focus window first
                time.sleep(0.1)
                
                # Set text directly
                self.textbox.set_text(data)
                logging.info(f"set_text() successful: {data}")
                time.sleep(0.2)
                
                # Press Enter
                self.textbox.type_keys('{ENTER}', pause=0.1)
                logging.info(f"Data input successful to '{self.textbox_auto_id}': {data}")
                
                # Wait for Shop-Flow to process data (increased wait time)
                time.sleep(1.0)  # Increased from 0.5 to 1.0 seconds
                
                # Check lblError popup after input
                self.check_lbl_error_popup()
            except Exception as e1:
                logging.error(f"set_text() method failed: {e1}, trying type_keys()...")
                # Method 2: type_keys
                try:
                    self.window.set_focus()
                    self.textbox.set_focus()
                    time.sleep(0.1)
                    self.textbox.type_keys(data + '{ENTER}', pause=0.1)
                    logging.info(f"type_keys() successful: {data}")
                    time.sleep(1.0)  # Increased wait time
                    self.check_lbl_error_popup()
                except Exception as e2:
                    logging.error(f"type_keys() also failed: {e2}")
        except Exception as e:
            logging.error(f"WinForms input error: {type(e).__name__} - {str(e)}")
            logging.error(f"Exception details: {repr(e)}")
            if "[WinError 5]" in str(e):
                messagebox.showerror("Error", "Please run as administrator")

    def send_ng_to_serial(self):
        try:
            # Send NG back to serial
            if self.serial_conn:
                bytes_written = self.serial_conn.write(b'NG\n')
                self.serial_conn.flush()  # Đảm bảo dữ liệu được gửi ngay
                logging.warning(f"⚠️ NG serial transmission successful ({bytes_written} bytes)")
            else:
                logging.error("❌ NG transmission failed - no serial connection")
        except Exception as e:
            logging.error(f"❌ NG transmission error: {type(e).__name__} - {e}")

    def send_ok_to_serial(self):
        try:
            # Send OK back to serial
            if self.serial_conn:
                bytes_written = self.serial_conn.write(b'OK\n')
                self.serial_conn.flush()  # Ensure data is sent immediately
                logging.info(f"✅ OK serial transmission successful ({bytes_written} bytes)")
            else:
                logging.error("❌ OK transmission failed - no serial connection")
        except Exception as e:
            logging.error(f"❌ OK transmission error: {type(e).__name__} - {e}")

    def click_reset_button(self):
        """Click the Reset button on Shop-Flow using keyboard shortcuts"""
        try:
            if not self.window:
                logging.error("❌ Shop-Flow window not initialized")
                return False
            
            # Focus on the Shop-Flow window first
            self.window.set_focus()
            time.sleep(0.2)
            
            # First shortcut: Alt+C
            logging.info("⌨️ Pressing Alt+C...")
            self.window.type_keys('%c')  # %c = Alt+C
            time.sleep(1.0)  # Wait 1 second
            
            # Second shortcut: Alt+R
            logging.info("⌨️ Pressing Alt+R...")
            self.window.type_keys('%r')  # %r = Alt+R
            
            logging.info("✅ Reset completed successfully (Alt+C → Alt+R)")
            time.sleep(0.5)
            return True
            
        except Exception as e:
            logging.error(f"❌ Reset button click error: {type(e).__name__} - {e}")
            return False

    def check_lbl_error_popup(self):
        """Check if lblError popup or NG dialog is visible and send OK/NG accordingly"""
        try:
            is_ng = False
            
            # Method 1: Check for lblError element
            try:
                lbl_error = self.window.child_window(auto_id='lblError')
                if lbl_error.exists() and lbl_error.is_visible():
                    logging.warning("lblError popup detected - sending NG")
                    is_ng = True
            except:
                pass
            
            # Method 2: Check for any visible window/dialog with "NG" text
            if not is_ng:
                try:
                    # Get all child windows
                    children = self.window.children()
                    for child in children:
                        try:
                            text = child.window_text()
                            # Check if control has "NG" text and is visible
                            if text and "NG" in text and child.is_visible():
                                # Check if it's a large control (likely the NG popup)
                                rect = child.rectangle()
                                width = rect.width()
                                height = rect.height()
                                # If control is large enough (e.g., > 200x200), it's probably the NG popup
                                if width > 200 and height > 200:
                                    logging.warning(f"NG popup detected (text='{text}', size={width}x{height}) - sending NG")
                                    is_ng = True
                                    break
                        except:
                            continue
                except:
                    pass
            
            # Method 3: Check for red color in large area (NG screen is red)
            if not is_ng:
                try:
                    # Try to find controls with "NG" in class name or control type
                    ng_controls = self.window.descendants(control_type="Window")
                    for ctrl in ng_controls:
                        try:
                            if ctrl.is_visible():
                                text = ctrl.window_text()
                                if "NG" in str(text):
                                    logging.warning(f"NG control detected: {text}")
                                    is_ng = True
                                    break
                        except:
                            continue
                except:
                    pass
            
            # Send result
            if is_ng:
                self.send_ng_to_serial()
            else:
                logging.info("No NG indicators found - sending OK")
                self.send_ok_to_serial()
                
        except Exception as e:
            # If can't determine, check the error message for common NG indicators
            error_msg = str(e).lower()
            if "ng" in error_msg or "error" in error_msg or "fail" in error_msg:
                logging.warning(f"Exception suggests NG: {e}")
                self.send_ng_to_serial()
            else:
                logging.info(f"Cannot determine status (assuming OK): {e}")
                self.send_ok_to_serial()

    def list_running_windows(self):
        """List running windows"""
        try:
            from pywinauto import Desktop
            desktop = Desktop(backend=self.backend)
            windows = desktop.windows()
            
            logging.info("Running windows:")
            for window in windows:
                try:
                    title = window.window_text()
                    if title and len(title.strip()) > 0:  # Exclude empty titles
                        logging.info(f"  - '{title}'")
                        # Check for partial match
                        if "shop" in title.lower() or "flow" in title.lower() or "Indonesia" in title.lower():
                            logging.info(f"    ^ Potential match for target app!")
                except:
                    continue
        except Exception as e:
            logging.error(f"Failed to list windows: {e}")
            
    def find_window_by_partial_title(self, keywords):
        """Find window by partial title"""
        try:
            from pywinauto import Desktop
            desktop = Desktop(backend=self.backend)
            windows = desktop.windows()
            
            for window in windows:
                try:
                    title = window.window_text()
                    if title:
                        for keyword in keywords:
                            if keyword.lower() in title.lower():
                                logging.info(f"Found potential window: '{title}'")
                                return window
                except:
                    continue
        except Exception as e:
            logging.error(f"Failed to search by partial title: {e}")
        return None

    def start(self):
        # For testing purposes, allow to continue even if serial connection fails
        serial_connected = self.connect_serial()
        if not serial_connected:
            logging.warning("Serial connection failed, but continuing to test WinForms connection...")
        
        # Display running windows list
        self.list_running_windows()
        
        try:
            logging.info(f"Trying to connect to application with title: '{self.target_app_title}'")
            
            # Try multiple connection methods
            try:
                # Method 1: Connect by exact title
                self.app = pywinauto.Application(backend=self.backend).connect(title=self.target_app_title)
                logging.info("Connected using exact title match")
            except:
                try:
                    # Method 2: Connect by partial title
                    self.app = pywinauto.Application(backend=self.backend).connect(title_re=".*Shop-Flow.*Indonesia.*")
                    logging.info("Connected using partial title match")
                except:
                    # Method 3: Try to find by process name (common for .NET apps)
                    try:
                        self.app = pywinauto.Application(backend=self.backend).connect(path="MIGHTY.ASFC.ITMPACK.exe")
                        logging.info("Connected using process name")
                    except:
                        # Method 4: Connect to any window with "Shop" or "Indonesia" in title
                        import pywinauto
                        desktop = pywinauto.Desktop(backend=self.backend)
                        for window in desktop.windows():
                            try:
                                title = window.window_text()
                                if title and ("shop" in title.lower() or "Indonesia" in title.lower()):
                                    self.app = pywinauto.Application(backend=self.backend).connect(handle=window.handle)
                                    logging.info(f"Connected to window: '{title}'")
                                    break
                            except:
                                continue
                        else:
                            raise Exception("Could not connect to target application using any method")
            
            # List all windows in the app
            windows = self.app.windows()
            logging.info(f"Found {len(windows)} windows in the application:")
            for idx, win in enumerate(windows):
                try:
                    logging.info(f"  Window {idx}: {win.window_text()}, class: {win.class_name()}, type: {type(win).__name__}")
                except Exception as we:
                    logging.error(f"  Window {idx}: Error - {we}")
            
            # Find main window (not dialog)
            self.window = None
            for win in windows:
                try:
                    # Find window with "Shop-Flow" in title
                    if 'Shop-Flow' in win.window_text():
                        self.window = win
                        logging.info(f"Selected main window: {win.window_text()}")
                        break
                except:
                    continue
            
            if not self.window:
                raise Exception("Could not find main application window")
            
            self.window.set_focus()
            
            logging.info(f"Looking for textbox with auto_id: '{self.textbox_auto_id}'")
            
            # DialogWrapper uses different API - convert to proper wrapper
            try:
                from pywinauto.controls.hwndwrapper import HwndWrapper
                # Recreate window wrapper with top_level_only=False to find children
                self.window = self.app.window(title_re=".*Shop-Flow.*", top_level_only=False)
                self.textbox = self.window.child_window(auto_id=self.textbox_auto_id, found_index=0)
            except Exception as wrap_err:
                logging.error(f"Failed to get textbox using child_window: {wrap_err}")
                # Fallback: Find directly from app
                try:
                    self.textbox = self.app.window(auto_id=self.textbox_auto_id, found_index=0)
                    logging.info(f"Found textbox using app.window() directly")
                except Exception as direct_err:
                    logging.error(f"Failed to find textbox directly: {direct_err}")
                    self.textbox = None
            
            if not self.textbox:
                logging.error(f"Textbox not found: auto_id '{self.textbox_auto_id}'")
                # List available controls
                try:
                    logging.info("Available controls in the window:")
                    for control in self.window.children():
                        try:
                            auto_id = control.automation_id()
                            class_name = control.class_name()
                            logging.info(f"  - auto_id: '{auto_id}', class: '{class_name}'")
                        except:
                            continue
                except Exception as ctrl_e:
                    logging.error(f"Failed to list controls: {ctrl_e}")
                return
            logging.info("WinForms app connection and textbox discovery successful")
        except Exception as e:
            logging.error(f"WinForms app connection failed: {e}")
            logging.error(f"Make sure the application '{self.target_app_title}' is running")
            return
        self.running = True
        self.thread = threading.Thread(target=self.read_serial_data)
        self.thread.daemon = True
        self.thread.start()
        logging.info("Background process started")

    def stop(self):
        self.running = False
        if self.serial_conn:
            self.serial_conn.close()
        logging.info("Process stopped")

if __name__ == "__main__":
    print(f"Process ID: {os.getpid()}")
    handler = SerialToWinForms()
    handler.start()

    # Don't exit even if no serial connection - can still test WinForms connection
    if not handler.serial_conn:
        logging.warning("Still waiting...")
    else:
        logging.info("Serial connection successful!")

    # Create system tray icon
    image = Image.new('RGB', (64, 64), color='blue')
    draw = ImageDraw.Draw(image)
    draw.ellipse([16, 16, 48, 48], fill='white')
    icon = pystray.Icon("Serial to WinForms", image, "Serial to WinForms")

    def quit_action(icon, item):
        icon.stop()
        handler.stop()
        os._exit(0)

    icon.menu = pystray.Menu(pystray.MenuItem("Quit", quit_action))
    icon.run_detached()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        icon.stop()
        handler.stop()