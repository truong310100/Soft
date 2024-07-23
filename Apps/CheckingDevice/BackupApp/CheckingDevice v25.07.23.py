#-*- mã hóa: utf-8 -*-
import webbrowser
import subprocess
from bs4 import BeautifulSoup
import os
from openpyxl import load_workbook
import datetime
import msvcrt

os.system('powercfg /batteryreport')
print('')

# Mở task manager
# os.system('start taskmgr')

# _________________________________________________________________________________________________________________

# Đọc nội dung file battery-report.html
with open('battery-report.html', 'r') as f:
    html_doc = f.read()

# Phân tích cú pháp HTML bằng BeautifulSoup
soup = BeautifulSoup(html_doc, 'html.parser')

# Tìm các thẻ td chứa giá trị "DESIGN CAPACITY" và "FULL CHARGE CAPACITY"
td_design_capacity = soup.find('td', string='DESIGN CAPACITY')
td_full_charge_capacity = soup.find('td', string='FULL CHARGE CAPACITY')

# Lấy giá trị của các cột "DESIGN CAPACITY" và "FULL CHARGE CAPACITY"
design_capacity = int(td_design_capacity.find_next_sibling('td').text.strip().replace(',', '').replace(' mWh', ''))
full_charge_capacity = int(td_full_charge_capacity.find_next_sibling('td').text.strip().replace(',', '').replace(' mWh', ''))

# Tính phần còn lại của pin
remaining_capacity = full_charge_capacity / design_capacity * 100

# _________________________________________________________________________________________________________________

# Lấy thông tin máy

# Device Name
hostname = os.popen('hostname').read().strip()
print('Device name:', hostname)

# Get System Model
result = subprocess.run(['wmic', 'csproduct', 'get', 'name'], capture_output=True, text=True)
if result.stderr:
    print(f"Error: {result.stderr.strip()}")
else:
    system_model = result.stdout.strip().split('\n')[2]
    print(f"System Model: {system_model}")

# Serial Number
result = subprocess.run(['wmic', 'bios', 'get', 'serialnumber'], capture_output=True, text=True)
if result.stderr:
    print(f"Error: {result.stderr.strip()}")
else:
    serial = result.stdout.strip().split('\n')[2]
    print(f"Serial Number: {serial}")
print('===========================================================================================================')

# CPU
result = subprocess.run(['wmic', 'cpu', 'get', 'name'], capture_output=True, text=True)
if result.stderr:
    print(f"Error: {result.stderr.strip()}")
else:
    CPU_Name = result.stdout.strip().split('\n')[2]
    print(f"CPU: {CPU_Name}")

# RAM
capacities = subprocess.run(['wmic', 'MemoryChip', 'get', 'Capacity'], capture_output=True, text=True).stdout.strip().split('\n')[2:]
speeds = subprocess.run(['wmic', 'MemoryChip', 'get', 'speed'], capture_output=True, text=True).stdout.strip().split('\n')[2:]
manufacturers = subprocess.run(['wmic', 'MemoryChip', 'get', 'manufacturer'], capture_output=True, text=True).stdout.strip().split('\n')[2:]

ram_info = [(cap.strip(), speed.strip(), man.strip()) for cap, speed, man in zip(capacities, speeds, manufacturers) if cap]
rams = []
for i, (capacity, speed, manufacturer) in enumerate(ram_info):
    Ram_OS = (f"RAM {i + 1}: {capacity:.1}GB {speed} {manufacturer}")
    rams.append(Ram_OS)
    print(Ram_OS)
Ram_OS = '\n'.join(rams)

# Disk
cmd = 'Get-PhysicalDisk | Select-Object MediaType, Manufacturer, Model, Size'
result = subprocess.run(['powershell', '-Command', cmd], capture_output=True, text=True)

if result.stderr:
    print(f"Error: {result.stderr.strip()}")
else:
    output = result.stdout.strip().splitlines()[2:]
    print('Disk:')
    disks = []
    i = 0
    for line in output:
        i += 1 
        parts = line.strip().split()
        media_type = parts[0]
        manufacturer = parts[1]
        model = ' '.join(parts[2:-1])
        size_bytes = int(parts[-1])
        size_gb = size_bytes / 2**30
        Disk = (f"Media {i}: {media_type}, Manufacturer: {manufacturer}, Model: {model}, Size: {size_gb:.2f} GB")
        disks.append(Disk)
        print(Disk)
    Disk = '\n'.join(disks)

print('===========================================================================================================')

# In ra kết quả Pin
print('Design capacity:', design_capacity)
print('Full charge capacity:', full_charge_capacity)
print('Pin: {:.2f}%'.format(remaining_capacity))

# Mở ứng dụng CrystalDiskInfo
os.system('start CrytalDiskInfo\DiskInfo64.exe')

# Web Check Driver
webbrowser.open('https://www.youtube.com/watch?v=d9LdhipeZ40')
webbrowser.open('https://www.youtube.com/watch?v=BbsOatt-3mA')
webbrowser.open('https://en.key-test.ru/')
webbrowser.open('https://www.onlinemictest.com/webcam-test/')
webbrowser.open('https://www.onlinemictest.com/')
webbrowser.open('battery-report.html')

print('===========================================================================================================')
print('')
print('__________Developed by HHO-TruongNL__________')
# print('__________Input Information Device___________')
print('')
input_username=input('Họ & tên người dùng:')
text_note_1=input('Tình trạng thân máy:')
text_note_2=input('Kết nối khác:')
text_note_3=input('Tuổi thọ ổ cứng:')

os.remove('battery-report.html')
# _________________________________________________________________________________________________________________
# Xuất Excel

# Mở file Excel
workbook = load_workbook(filename='FileExport/BanGiao.xlsx')

# Lấy sheet cần ghi dữ liệu
sheet = workbook['BienBan']

# Ghi dữ liệu vào ô
while True:
    print(" Bạn là người giao máy = Nhấn phím 1 \n Bạn là người nhận máy = Nhấn phím 2 \n Mời bạn nhấn phím 1 hoặc 2")
    choice = msvcrt.getch().decode()

    if choice == '1':
        choice = "Người giao:"
        BienBan = "Giao"
        sheet['A10'] = choice
        print(choice)
        valid_choice = True
        break
    elif choice == '2':
        choice = "Người nhận:"
        BienBan = "Nhan"
        sheet['A10'] = choice
        print(choice)
        valid_choice = True
        break

sheet['B12'] = input_username
sheet['D14'] = hostname
sheet['E17'] = serial
sheet['B17'] = system_model
sheet['c17'] = f"{CPU_Name}\n{Ram_OS}\n{Disk}"

text_note_4 = ('Pin còn lại: {:.2f}%'.format(remaining_capacity))
note = (f"Tình trạng thân máy: {text_note_1}\nKết nối khác: {text_note_2}\nTuổi thọ ổ cứng: {text_note_3}\n{text_note_4}")
sheet['F17'] = note

# Đặt tên file
date = datetime.date.today()
date_string = date.strftime("%d%m%Y")

# Lưu lại file Excel
filename = f'FileExport/BB-{BienBan}-{input_username}-{date_string}.xlsx'
workbook.save(filename)
os.system("start "" FileExport")
subprocess.run(['start', '', filename], shell=True)

# Mở Sharepoint và lấy file kết quả excel
webbrowser.open('https://hunghaugroup.sharepoint.com/:f:/g/HHH.HungHauHoldings/Es1sMfxQCtJMvZjz8MH6trMB9do2xRhMbPQbPJSXy1pjww?e=BReDQB')