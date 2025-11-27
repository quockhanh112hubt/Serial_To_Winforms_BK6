#!/usr/bin/env python3
"""
Test Serial Sender - Mô phỏng machine gửi dữ liệu vào COM port
Chạy file này để test serial_to_winforms_bk6 đang chạy ở chế độ chờ
"""

import serial
import serial.tools.list_ports
import time
import random
import string
import sys
import threading
import logging

class SerialTestSender:
    def __init__(self):
        self.com_port = None
        self.baudrate = 9600  # PHẢI KHỚP VỚI CONFIG.TXT!
        self.serial_conn = None
        self.running = False
        
        # Thiết lập logging
        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )
        
    def list_available_ports(self):
        """Liệt kê các COM port có sẵn"""
        ports = serial.tools.list_ports.comports()
        available_ports = []
        
        print("=== Available COM Ports ===")
        for i, port in enumerate(ports, 1):
            print(f"{i}. {port.device} - {port.description}")
            available_ports.append(port.device)
            
        return available_ports
        
    def select_port(self):
        """Cho phép user chọn COM port"""
        available_ports = self.list_available_ports()
        
        if not available_ports:
            print("❌ Không tìm thấy COM port nào!")
            return False
            
        try:
            choice = input(f"\nChọn COM port (1-{len(available_ports)}) hoặc nhập trực tiếp (ví dụ: COM5): ")
            
            # Nếu nhập số
            if choice.isdigit():
                idx = int(choice) - 1
                if 0 <= idx < len(available_ports):
                    self.com_port = available_ports[idx]
                else:
                    print("❌ Lựa chọn không hợp lệ!")
                    return False
            # Nếu nhập trực tiếp COM port
            elif choice.upper().startswith('COM'):
                self.com_port = choice.upper()
            else:
                print("❌ Format không đúng!")
                return False
                
            print(f"✅ Đã chọn: {self.com_port}")
            return True
            
        except KeyboardInterrupt:
            print("\n❌ Hủy bởi người dùng")
            return False
            
    def connect_to_port(self):
        """Kết nối tới COM port"""
        try:
            print(f"🔌 Đang kết nối tới {self.com_port}...")
            self.serial_conn = serial.Serial(
                port=self.com_port,
                baudrate=self.baudrate,
                timeout=1,
                write_timeout=1
            )
            print(f"✅ Kết nối thành công!")
            return True
            
        except Exception as e:
            print(f"❌ Kết nối thất bại: {e}")
            return False
            
    def generate_test_data(self):
        """Tạo dữ liệu test theo format STX...ETX"""
        
        # Chỉ gửi Pattern 2: Weight data
        # data = "01462008114854321I44HK24AZCK5240DKR00009.00;01462008114854321I44HK24AYCIY240DKR00009.00;01462008114854321I44HK24AXUOI240DKR00009.00;01462008114854321I44HK24AWRIQ240DKR00009.00;01462008114854321I44HK24AV2OP240DKR00009.00;01462008114854321I44HK24AUXZ9240DKR00009.00;01462008114854321I44HK24AR6JW240DKR00009.00;01462008114854321I44HK24AQUSI240DKR00009.00;01462008114854321I44HK24AN2FY240DKR00009.00;01462008114854321I44HK24AKVRB240DKR00009.00;01462008114854321I44HK24AIHHN240DKR00009.00;01462008114854321I44HK24AH0HV240DKR00009.00;01462008114854321I44HK24AE0K6240DKR00009.00;01462008114854321I44HK24ADKWJ240DKR00009.00;01462008114854321I44HK24ACM9A240DKR00009.00;01462008114854321I44HK24A5CUG240DKR00009.00;01462008114854321I44HK24A58WA240DKR00009.00;01462008114854321T44HJ20AI0OH240DKR00009.00;01462008114854321T44HJ21A23L1240DKR00009.00;01462008114854321T44HJ21ADOY0240DKR00009.00"
        data = "RESET"
        # Thêm STX và ETX
        formatted_data = f"STX{data}ETX"
        return formatted_data, data
        
    def send_single_data(self):
        """Gửi một lần dữ liệu"""
        if not self.serial_conn:
            print("❌ Chưa kết nối COM port!")
            return False
            
        try:
            full_data, core_data = self.generate_test_data()
            
            print(f"📤 Gửi: {full_data}")
            print(f"   → Core data: {core_data}")
            
            # Gửi dữ liệu
            self.serial_conn.write((full_data + '\n').encode('utf-8'))
            self.serial_conn.flush()
            
            print("✅ Gửi thành công!")
            
            # Đợi phản hồi
            print("⏳ Đang chờ phản hồi (OK/NG)...")
            start_time = time.time()
            
            while time.time() - start_time < 5:  # Timeout 5 giây
                if self.serial_conn.in_waiting > 0:
                    response = self.serial_conn.readline().decode('utf-8').strip()
                    if response:
                        print(f"📥 Phản hồi: {response}")
                        if response == "OK":
                            print("✅ Kết quả: THÀNH CÔNG")
                        elif response == "NG":
                            print("❌ Kết quả: LỖI - Machine sẽ dừng!")
                        else:
                            print(f"⚠️ Phản hồi không xác định: {response}")
                        return True
                time.sleep(0.1)
                
            print("⏰ Timeout - Không nhận được phản hồi")
            return False
            
        except Exception as e:
            print(f"❌ Lỗi gửi dữ liệu: {e}")
            return False
            
    def auto_send_mode(self):
        """Chế độ gửi tự động"""
        print("\n=== Chế độ gửi tự động ===")
        try:
            interval = float(input("Nhập khoảng thời gian giữa các lần gửi (giây): "))
        except:
            interval = 3.0
            
        print(f"🚀 Bắt đầu gửi tự động mỗi {interval} giây")
        print("Nhấn Ctrl+C để dừng...")
        
        self.running = True
        count = 0
        
        try:
            while self.running:
                count += 1
                print(f"\n--- Lần gửi #{count} ---")
                
                success = self.send_single_data()
                if success:
                    print("Chờ khoảng thời gian tiếp theo...")
                else:
                    print("Gửi thất bại, tiếp tục...")
                    
                time.sleep(interval)
                
        except KeyboardInterrupt:
            print(f"\n🛑 Đã dừng sau {count} lần gửi")
            self.running = False
            
    def manual_send_mode(self):
        """Chế độ gửi thủ công"""
        print("\n=== Chế độ gửi thủ công ===")
        print("Nhấn Enter để gửi dữ liệu, 'q' để thoát")
        
        while True:
            user_input = input("\nNhấn Enter để gửi (hoặc 'q' để thoát): ").strip()
            
            if user_input.lower() == 'q':
                break
                
            self.send_single_data()
            
    def check_com_port_conflict(self):
        """Kiểm tra xung đột COM port với chương trình chính"""
        try:
            with open('config.txt', 'r') as f:
                content = f.read()
                if f"port = {self.com_port}" in content:
                    print(f"⚠️  CẢNH BÁO: {self.com_port} đang được sử dụng bởi serial_to_winforms_bk6!")
                    print("   Điều này sẽ gây xung đột.")
                    print("\n💡 GIẢI PHÁP:")
                    print("1. Dừng serial_to_winforms_bk6 trước")
                    print("2. Hoặc sử dụng COM port khác") 
                    print("3. Hoặc sử dụng Virtual COM Port Pair")
                    
                    choice = input("\nBạn có muốn tiếp tục không? (y/n): ").lower()
                    return choice == 'y'
        except:
            pass
        return True

    def run(self):
        """Chạy chương trình chính"""
        print("=" * 50)
        print("🧪 SERIAL TEST SENDER")
        print("Mô phỏng machine gửi dữ liệu vào COM port")
        print("=" * 50)
        
        print("\n📋 LưU Ý QUAN TRỌNG:")
        print("- Test sender này GỬI dữ liệu VÀO COM port")
        print("- serial_to_winforms_bk6 ĐỌNG dữ liệu TỪ COM port") 
        print("- Cần dùng 2 COM port khác nhau hoặc Virtual COM Pair")
        print("- Hoặc dừng serial_to_winforms_bk6 để test riêng")
        
        # Bước 1: Chọn COM port
        if not self.select_port():
            return
            
        # Bước 1.5: Kiểm tra xung đột COM port
        if not self.check_com_port_conflict():
            return
            
        # Bước 2: Kết nối
        if not self.connect_to_port():
            return
            
        # Bước 3: Chọn chế độ test
        print("\n=== Chọn chế độ test ===")
        print("1. Gửi thủ công (Manual)")
        print("2. Gửi tự động (Auto)")
        print("3. Gửi một lần và thoát")
        
        try:
            mode = input("Chọn chế độ (1-3): ").strip()
            
            if mode == "1":
                self.manual_send_mode()
            elif mode == "2":
                self.auto_send_mode()
            elif mode == "3":
                self.send_single_data()
            else:
                print("❌ Chế độ không hợp lệ!")
                
        except KeyboardInterrupt:
            print("\n🛑 Chương trình bị ngắt")
            
        finally:
            if self.serial_conn:
                self.serial_conn.close()
                print("🔌 Đã đóng kết nối COM port")

if __name__ == "__main__":
    print("🚀 Khởi động Serial Test Sender...")
    print("📋 Đảm bảo serial_to_winforms_bk6 đang chạy ở chế độ chờ")
    
    sender = SerialTestSender()
    sender.run()
    
    print("👋 Kết thúc chương trình")