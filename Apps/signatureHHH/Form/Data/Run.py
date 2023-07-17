from tkinter import *
from docxtpl import DocxTemplate
import os

class App(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        # Tạo một menu dropdown
        self.var = StringVar(self)
        self.var.set("Khối")
        self.dropdown = OptionMenu(self, self.var, "HHO", "HHA", "HA1", "HHE", "HHD", "HHF", "HCM", "VHU", "Vạn Tường", "OSV")
        self.dropdown.grid(row=0, column=1)

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

        # Tạo nút để tạo tệp Word
        self.button = Button(self, text="Tạo", command=self.create_word_file).grid(row=6, column=1)
        self.master.bind('<Return>', lambda event: self.create_word_file())

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

        # Tạo tệp Word
        if self.var.get() == "HHO":
            word = DocxTemplate(r'Form\HHO.docx')
        elif self.var.get() == "HHA":
            word = DocxTemplate(r'Form\HHA.docx')
        elif self.var.get() == "HA1":
            word = DocxTemplate(r'Form\HA1.docx')
        elif self.var.get() == "HHE":
            word = DocxTemplate(r'Form\HHE.docx')
        elif self.var.get() == "HHD":
            word = DocxTemplate(r'Form\HHD.docx')
        elif self.var.get() == "HHF":
            word = DocxTemplate(r'Form\HHF.docx')
        elif self.var.get() == "HCM":
            word = DocxTemplate(r'Form\HCM.docx')
        elif self.var.get() == "VHU":
            word = DocxTemplate(r'Form\VHU.docx')
        elif self.var.get() == "Vạn Tường":
            word = DocxTemplate(r'Form\VanTuong.docx')
        elif self.var.get() == "OSV":
            word = DocxTemplate(r'Form\OSV.docx')
        word.render(context)
        # word.save(os.path.join('output', f"{self.var.get()}-{context['HovaTen']}.docx"))
        # os.startfile(os.path.join('output', f"{self.var.get()}-{context['HovaTen']}.docx"))
        # print("Tạo thành công")

        output_folder = 'Output'
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        file_name = f"{self.var.get()}-{context['HovaTen']}.docx"
        file_path = os.path.join(output_folder, file_name)

        word.save(file_path)
        os.startfile(file_path)

root = Tk()
root.geometry('400x200')
root.title("Signature HHH - v27.06.23")
app = App(master=root)
app.mainloop()