from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from docxtpl import DocxTemplate
import os
import win32com.client as win32
from PIL import Image, ImageTk
import requests
from io import BytesIO

class App(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        # Thêm nút để chọn file mẫu Word
        self.label_file = Label(self, text="Chọn file Word:")
        self.label_file.grid(row=0, column=0)
        self.button_file = Button(self, text="Chọn File", command=self.select_word_template)
        self.button_file.grid(row=0, column=2)
        self.word_template_entry = Entry(self, width=30)
        self.word_template_entry.grid(row=0, column=1)

        self.label1 = Label(self, text="Họ và tên:")
        self.label1.grid(row=1, column=0)
        self.entry1 = Entry(self, width=30)
        self.entry1.grid(row=1, column=1)
        self.entry1.bind("<KeyRelease>", self.format_name)

        self.label2 = Label(self, text="Chức Vụ:")
        self.label2.grid(row=2, column=0)
        self.entry2 = Entry(self, width=30)
        self.entry2.grid(row=2, column=1)

        self.label3 = Label(self, text="Phòng Ban:")
        self.label3.grid(row=3, column=0)
        self.entry3 = Entry(self, width=30)
        self.entry3.grid(row=3, column=1)

        self.label4 = Label(self, text="Mail:")
        self.label4.grid(row=4, column=0)
        self.entry4 = Entry(self, width=30)
        self.entry4.grid(row=4, column=1)
        self.entry4.bind("<KeyRelease>", self.format_mail)

        self.label5 = Label(self, text="Sđt (+84):")
        self.label5.grid(row=5, column=0)
        self.entry5 = Entry(self, width=30)
        self.entry5.grid(row=5, column=1)
        self.entry5.bind("<KeyRelease>", self.format_phone_number)

        # Frame để chứa hai nút
        self.button_frame = Frame(self)
        self.button_frame.grid(row=6, column=1, pady=10, sticky=W+E)

        # Nút để tạo tệp Word
        self.button_create = Button(self.button_frame, text="Tạo Word", width=10, command=self.create_word_file)
        self.button_create.grid(row=0, column=0, sticky=W+E)

        # Nút để gửi mail
        self.button_send = Button(self.button_frame, text="Gửi Mail", width=10, command=self.send_mail)
        self.button_send.grid(row=0, column=1, sticky=W+E)

        self.master.bind('<Return>', lambda event: self.create_word_file())
        self.master.bind('<Return>', lambda event: self.send_mail())

    def select_word_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if file_path:
            self.word_template_entry.delete(0, END)
            self.word_template_entry.insert(0, file_path)

    def format_name(self, event):
        s = self.entry1.get()
        s = s.title()
        self.entry1.delete(0, END)
        self.entry1.insert(0, s)

    def format_mail(self, event):
        s = self.entry4.get()
        s = s.lower()
        self.entry4.delete(0, END)
        self.entry4.insert(0, s)

    def format_phone_number(self, event):
        s = self.entry5.get().replace(" ", "")
        # Xóa số 0 ở đầu chuỗi
        if s.startswith("+84"):
            s = s[3:]
        if s.startswith("0"):
            s = s[1:]
        if len(s) > 3:
            s = s[:2] + " " + s[2:]
        if len(s) > 7:
            s = s[:7] + " " + s[7:]
        self.entry5.delete(0, END)
        self.entry5.insert(0, s)

    def create_word_file(self):
        # Lấy dữ liệu từ các entry
        context = {
            'HovaTen': self.entry1.get(),
            'ChucVu': self.entry2.get(),
            'PhongBan': self.entry3.get(),
            'Mail': self.entry4.get(),
            'Phone': self.entry5.get(),
        }
        # Lấy đường dẫn tệp mẫu từ entry
        template_path = self.word_template_entry.get()
        if not template_path:
            messagebox.showinfo("Thông báo","Vui lòng chọn file mẫu Word.")
            return
        # Tạo tệp Word
        word = DocxTemplate(template_path)
        word.render(context)
        current_directory = os.getcwd()
        output_folder = os.path.join(current_directory, 'output')
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        file_path = os.path.join(output_folder, f"{context['Mail']}.docx")
        word.save(file_path)
        os.startfile(file_path)

    def send_mail(self):
        # Lấy dữ liệu từ các entry
        context = {
            'HovaTen': self.entry1.get(),
            'ChucVu': self.entry2.get(),
            'PhongBan': self.entry3.get(),
            'Mail': self.entry4.get(),
            'Phone': self.entry5.get(),
        }
        # Lấy đường dẫn tệp mẫu từ entry
        template_path = self.word_template_entry.get()
        if not template_path:
            messagebox.showinfo("Thông báo","Vui lòng chọn file mẫu Word.")
            return
        # Tạo tệp Word
        word = DocxTemplate(template_path)
        word.render(context)
        current_directory = os.getcwd()
        output_folder = os.path.join(current_directory, 'output')
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        email = context['Mail']
        file_path = os.path.join(output_folder, f"{context['Mail']}.docx")
        word.save(file_path)
        self.send_email_with_word(email, file_path)

    def send_email_with_word(self, email, file_path):
        print(f"Đang gửi email tới: {email}")
        print(f"Đính kèm tệp: {file_path}")
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = "Tệp Word đã tạo"
        mail.Body = 'Chữ ký mới'
        mail.Attachments.Add(file_path)
        mail.Send()
        messagebox.showinfo("Thông báo", "Email đã được gửi thành công!")

root = tk.Tk()
url = "https://hunghau.vn/wp-content/uploads/2022/02/logo-HHH-new.png"
response = requests.get(url)
image_data = response.content
image = Image.open(BytesIO(image_data))
root.iconphoto(False, ImageTk.PhotoImage(image))
# root.minsize(width=500, height=300)
root.title("Signature HHH - v23.06.24")
app = App(master=root)
app.mainloop()