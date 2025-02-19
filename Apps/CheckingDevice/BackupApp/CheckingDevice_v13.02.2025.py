import os
import time
import json
import ctypes
import datetime
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk
from bs4 import BeautifulSoup
from tkinter import messagebox
from openpyxl import load_workbook

cmd_pin = 'powercfg /batteryreport'
cmd_crytal = 'start CrytalDiskInfo\\DiskInfo64.exe'
cmd_report_pin = 'battery-report.html'
cmd_rename = 'start BackupApp\\Rename_PC.cmd'
web_check_screen = 'https://www.youtube.com/watch?v=d9LdhipeZ40'
web_check_speaker='https://mymictest.com/vi/speaker-test/'
web_check_keyboard='https://en.key-test.ru/'
web_check_camera='https://www.onlinemictest.com/webcam-test/'
web_check_micro='https://www.onlinemictest.com/'
web_sharepoint = 'https://hunghaugroup.sharepoint.com/:f:/g/HHH.HungHauHoldings/Es1sMfxQCtJMvZjz8MH6trMB9do2xRhMbPQbPJSXy1pjww?e=BReDQB'

def load_support_data():
    try:
        json_path = r"BackupApp\data_localhost.json"
        with open(json_path, "r", encoding="utf-8") as file:
            data = json.load(file)
            return {
                "names": [person["name"] for person in data], 
                "working_locations": [person["working_location"] for person in data] 
            }
    except Exception as e:
        print(f"Lỗi khi đọc file JSON: {e}")
        return {"names": [], "working_locations": []}
support_data = load_support_data()
support_names = support_data["names"]
support_working_locations = support_data["working_locations"]

def check_pin():
    os.system(cmd_crytal)
    webbrowser.open(cmd_report_pin)
    
def check_driver():
    urls = [
        web_check_screen,
        web_check_speaker,
        web_check_keyboard,
        web_check_camera,
        web_check_micro
    ]
    
    for url in urls:
        webbrowser.open(url)
        time.sleep(0.5)
        
os.system(cmd_pin)
# check_driver()
# check_pin()

def submit_form(): # Lay thong tin thiết bị
    input_user_Support = support_combobox.get()
    input_user_name = entry_name.get()
    input_department = entry_department.get()
    text_note_1 = entry_note1.get()
    text_note_2 = entry_note2.get()
    text_note_3 = entry_note3.get()
    type_bb = type_var.get()
    selected_support = support_combobox.get()
    if selected_support in support_names:
        index = support_names.index(selected_support)
        working_location = support_working_locations[index]
    else:
        working_location = "Không xác định"
        
    # Tạo báo cáo pin
    with open('battery-report.html', 'r') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    design_capacity = int(soup.find('td', string='DESIGN CAPACITY').find_next_sibling('td').text.strip().replace(',', '').replace(' mWh', ''))
    full_charge_capacity = int(soup.find('td', string='FULL CHARGE CAPACITY').find_next_sibling('td').text.strip().replace(',', '').replace(' mWh', ''))
    remaining_capacity = full_charge_capacity / design_capacity * 100

    # Lấy thông tin hệ thống
    hostname = entry_hostname.get().strip()  # Lấy giá trị từ ô nhập
    
    system_model = subprocess.run(['wmic', 'csproduct', 'get', 'name'], capture_output=True, text=True).stdout.strip().split('\n')[2]
    serial = subprocess.run(['wmic', 'bios', 'get', 'serialnumber'], capture_output=True, text=True).stdout.strip().split('\n')[2]
    CPU_Name = subprocess.run(['wmic', 'cpu', 'get', 'name'], capture_output=True, text=True).stdout.strip().split('\n')[2]

    # Ghi vào file Excel
    workbook = load_workbook('BackupApp/BanGiao.xlsx')
    sheet = workbook['BienBan']
    sheet['B10'] = input_user_Support
    sheet['D10'] = working_location
    sheet['D12'] = input_department
    sheet['A10'], sheet['B12'], sheet['D14'] = ('Người giao:' if type_bb == 'Giao' else 'Người nhận:'), input_user_name, hostname
    sheet['E17'], sheet['B17'], sheet['C17'] = serial, system_model, CPU_Name
    sheet['F17'] = f"Tình trạng thân máy: {text_note_1}\nKết nối khác: {text_note_2}\nTuổi thọ ổ cứng: {text_note_3}\nPin còn lại: {remaining_capacity:.2f}%"

    # Lưu và mở file
    if not os.path.exists('Export'):
        os.makedirs('Export')
    filename = f'Export/BB-{type_bb}-{input_user_name}-{datetime.date.today().strftime("%d%m%Y")}.xlsx'
    workbook.save(filename)
    os.system(f'start "" "{filename}"')

    # Mở trang web kiểm tra
    os.system(web_sharepoint)
    os.remove(cmd_report_pin)

def rename_pc():
    os.system(cmd_rename)
        
# Giao diện tkinter
root = tk.Tk()
icon = tk.PhotoImage(file='BackupApp/logo-HHH.png')
root.tk.call('wm', 'iconphoto', root._w, icon)
root.title("Checking Device")
root.configure(bg="#f0f0f0")

# Tạo style cho các widget
style = ttk.Style()
style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
style.configure("TButton", font=("Arial", 10), padding=5)
style.configure("TEntry", font=("Arial", 10), padding=5)
style.configure("TCombobox", font=("Arial", 10), padding=5)

# Họ & tên người hỗ trợ (sử dụng Combobox)
ttk.Label(root, text="Họ & tên người hỗ trợ:", style="TLabel").grid(row=0, column=0, padx=10, pady=5, sticky="w")
support_var = tk.StringVar()
support_combobox = ttk.Combobox(root, textvariable=support_var, values=support_names, style="TCombobox")
support_combobox.grid(row=0, column=1, padx=10, pady=5, sticky="ew", ipadx=100)
support_combobox.current(0)  # Chọn mặc định giá trị đầu tiên

# Nút kiểm tra pin
ttk.Button(root, text="Kiểm tra pin", command=check_pin, style="TButton").grid(row=0, column=2, padx=10, pady=5)

# Nút kiểm tra kết nối
ttk.Button(root, text="Kiểm tra kết nối", command=check_driver, style="TButton").grid(row=1, column=2, padx=10, pady=5)

# Họ & tên người dùng
ttk.Label(root, text="Họ & tên người dùng:", style="TLabel").grid(row=1, column=0, padx=10, pady=5, sticky="w")
entry_name = ttk.Entry(root, style="TEntry")
entry_name.grid(row=1, column=1, padx=10, pady=5, sticky="ew", ipadx=100)

# Chức vụ - Phòng ban người dùng
ttk.Label(root, text="Chức vụ - Phòng ban người dùng:", style="TLabel").grid(row=2, column=0, padx=10, pady=5, sticky="w")
entry_department = ttk.Entry(root, style="TEntry")
entry_department.grid(row=2, column=1, padx=10, pady=5, sticky="ew", ipadx=100)

# Tên máy người dùng
ttk.Label(root, text="Tên máy người dùng:", style="TLabel").grid(row=3, column=0, padx=10, pady=5, sticky="w")
default_hostname = os.popen('hostname').read().strip()
entry_hostname = ttk.Entry(root, style="TEntry")
entry_hostname.insert(0, default_hostname)  # Gán giá trị mặc định
entry_hostname.grid(row=3, column=1, padx=10, pady=5, sticky="ew", ipadx=100)

# Nút đổi tên máy
ttk.Button(root, text="Đổi tên máy", command=rename_pc, style="TButton").grid(row=3, column=2, padx=10, pady=5)

# Tình trạng thân máy
ttk.Label(root, text="Tình trạng thân máy:", style="TLabel").grid(row=4, column=0, padx=10, pady=5, sticky="w")
entry_note1 = ttk.Entry(root, style="TEntry")
entry_note1.grid(row=4, column=1, padx=10, pady=5, sticky="ew", ipadx=100)

# Kết nối khác
ttk.Label(root, text="Kết nối khác:", style="TLabel").grid(row=5, column=0, padx=10, pady=5, sticky="w")
entry_note2 = ttk.Entry(root, style="TEntry")
entry_note2.grid(row=5, column=1, padx=10, pady=5, sticky="ew", ipadx=100)

# Tuổi thọ ổ cứng
ttk.Label(root, text="Tuổi thọ ổ cứng:", style="TLabel").grid(row=6, column=0, padx=10, pady=5, sticky="w")
entry_note3 = ttk.Entry(root, style="TEntry")
entry_note3.grid(row=6, column=1, padx=10, pady=5, sticky="ew", ipadx=100)

# Loại biên bản
ttk.Label(root, text="Loại biên bản:", style="TLabel").grid(row=7, column=0, padx=10, pady=5, sticky="w")
type_var = tk.StringVar()
type_select = ttk.Combobox(root, textvariable=type_var, values=["Giao", "Nhận"], style="TCombobox")
type_select.grid(row=7, column=1, padx=10, pady=5, sticky="ew", ipadx=100)
type_select.current(0)

# Nút gửi
ttk.Button(root, text="Tạo", command=submit_form, style="TButton").grid(row=8, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

root.columnconfigure(1, weight=1) # Căn chỉnh các cột
root.mainloop()