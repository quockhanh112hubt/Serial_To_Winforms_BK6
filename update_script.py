import os
import sys
import zipfile
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import time
import threading
from ftplib import FTP

# Đường dẫn tới thư mục chứa các file chương trình
PROGRAM_DIRECTORY = "C:\\Serial_to_MES"
CURRENT_VERSION_FILE = "C:\\Serial_to_MES\\version.txt"
MAIN_EXECUTABLE = os.path.join(PROGRAM_DIRECTORY, "SerialToWinForms.exe")
UPDATE_ZIP_PATH = os.path.join(PROGRAM_DIRECTORY, "update.zip")
VERSION_FLAG_FILE = os.path.join(PROGRAM_DIRECTORY, "version_flag.txt")

# Đường dẫn tới FTP Server
FTP_SERVER = "10.62.102.5"
FTP_USER = "update"
FTP_PASS = "update"
FTP_DIRECTORY = "KhanhDQ/Update_Program/Serial_to_MES/"

def get_current_version():
    if os.path.exists(VERSION_FLAG_FILE):
        with open(VERSION_FLAG_FILE, "r") as file:
            return file.read().strip()
    return "0.0.0"

def set_current_version(version):
    with open(VERSION_FLAG_FILE, "w") as file:
        file.write(version)

def update_version_file(new_version):
    with open(CURRENT_VERSION_FILE, "w") as file:
        file.write(new_version)

def get_latest_version():
    try:
        ftp = FTP(FTP_SERVER)
        ftp.login(FTP_USER, FTP_PASS)
        ftp.cwd(FTP_DIRECTORY)
        with open("latest_version.txt", "wb") as file:
            ftp.retrbinary("RETR version.txt", file.write)
        ftp.quit()
        with open("latest_version.txt", "r") as file:
            latest_version = file.read().strip()
        os.remove("latest_version.txt")
        return latest_version
    except Exception as e:
        print(f"Không thể lấy phiên bản mới nhất: {e}")
        return None

def download_update(progress_var):
    try:
        ftp = FTP(FTP_SERVER)
        ftp.login(FTP_USER, FTP_PASS)
        ftp.cwd(FTP_DIRECTORY)
        total_size = ftp.size("update.zip")
        block_size = 8192
        downloaded_size = 0
        with open(os.path.join(PROGRAM_DIRECTORY, "update.zip"), 'wb') as file:
            def callback(data):
                nonlocal downloaded_size
                file.write(data)
                downloaded_size += len(data)
                percent = round(downloaded_size * 100 / total_size)
                progress_var.set(percent)
                root.update_idletasks()
            ftp.retrbinary("RETR update.zip", callback, block_size)
        ftp.quit()
        print("Tải bản cập nhật thành công.")
    except Exception as e:
        print(f"Không thể tải bản cập nhật: {e}")
        messagebox.showerror("Lỗi", "Không thể tải bản cập nhật!. Kiểm tra kết nối mạng")
        close_window(root)
        sys.exit()

def apply_update():
    try:
        print("Đang giải nén file cập nhật...")
        with zipfile.ZipFile(UPDATE_ZIP_PATH, "r") as zip_ref:
            zip_ref.extractall(PROGRAM_DIRECTORY)
        zip_ref.close()
        os.remove(UPDATE_ZIP_PATH)
        print("Cập nhật hoàn tất.")
    except Exception as e:
        print(f"Lỗi khi thực hiện cập nhật: {e}")
        messagebox.showerror("Lỗi", "Không thể giản nén bản cập nhật!")
        close_window(root)
        sys.exit()

def restart_program(root):
    try:
        if os.path.exists(MAIN_EXECUTABLE):
            process = subprocess.Popen([MAIN_EXECUTABLE])
            print(f"Đã khởi chạy lại chương trình, PID: {process.pid}")
            close_window(root)
            sys.exit()
        else:
            print(f"Không tìm thấy tệp {MAIN_EXECUTABLE}")
            messagebox.showerror("Lỗi", f"Không tìm thấy tệp {MAIN_EXECUTABLE}")
    except Exception as e:
        print(f"Lỗi khi khởi động lại chương trình: {e}")
        messagebox.showerror("Lỗi", f"Lỗi khi khởi động lại chương trình: {e}")

def show_update_window(update_action):
    global root
    root = tk.Tk()
    root.title("Đang cập nhật")
    root.geometry("420x150")
    root.resizable(False, False) 
    
    label = tk.Label(root, text="Đang thực hiện cập nhật...", font=("Arial", 14))
    label.pack(pady=20)

    frame = tk.Frame(root)
    frame.pack(pady=10)

    progress_var = tk.DoubleVar()
    progress = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate", variable=progress_var, maximum=100)
    progress.pack(side="left", padx=(10, 0))
    
    percent_label = tk.Label(frame, text="0%", font=("Arial", 14))
    percent_label.pack(side="left", padx=(10, 0))
    def update_percent_label(*args):
        percent_label.config(text=f"{progress_var.get()}%")
    progress_var.trace("w", update_percent_label)


    def run_update():
        update_action(root, progress_var)
    
    threading.Thread(target=run_update).start()
    root.mainloop()

def close_window(root):
    root.destroy()

if __name__ == "__main__":
    def update_action(root, progress_var):

        if not os.path.exists(CURRENT_VERSION_FILE):
            messagebox.showerror("Lỗi", "Không tìm thấy tệp version.txt. Hãy kiểm tra lại!")
            close_window(root)
            sys.exit()

        download_update(progress_var)
        new_version = "0.0.0"
        try:
            with zipfile.ZipFile(UPDATE_ZIP_PATH, 'r') as zip_ref:
                new_version_file = zip_ref.open('version.txt')
                new_version = new_version_file.read().decode('utf-8').strip()
                new_version_file.close()
        except Exception as e:
            messagebox.showerror("Lỗi", "Không thể lấy thông tin phiên bản từ Server!")
            close_window(root)
            sys.exit()
        
        current_version = get_current_version()
        latest_version = get_latest_version()

        if new_version > current_version:
            time.sleep(2)
            apply_update()
            update_version_file(latest_version)
            set_current_version(new_version)
            restart_program(root)
        else:
            messagebox.showinfo("Thông báo", "Chương trình đã được cập nhật phiên bản mới nhất.")
            restart_program(root)
    
    show_update_window(update_action)
