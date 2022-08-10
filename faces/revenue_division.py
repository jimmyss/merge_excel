import csv
import ttkbootstrap as ttk
from faces import init_face
from tkinter import filedialog
from ttkbootstrap.constants import *

class revenue_division():

    def __init__(self, master):
        self.master = master
        self.master.config()

        #---------------变量--------------
        self.file_path=""

        #--------------月度分账收入界面------------
        self.division_face=ttk.Frame(self.master)
        self.division_face.grid()

        #--------------初始化月度分账收入界面组件----------
        #按钮组件
        self.choose_file=ttk.Button(self.division_face, text="选择文件", command=self.choose_file)#--------------选择文件按钮--------------
        self.choose_file.grid(row=0, column=0, pady=10, padx=10)
        self.back=ttk.Button(self.division_face, text="    返回    ", command=self.to_iniface)#--------------返回按钮-----------------
        self.back.grid(row=0, column=2, padx=10, pady=10)
        self.create_division=ttk.Button(self.division_face, text="制作收入分账", command=self.divide)#--------------制作收入分账按钮----------
        self.create_division.grid(row=1, column=2, pady=10, padx=10)
        self.create_division.grid_forget()

        #标签组件
        self.save_label=ttk.Label(self.division_face, text="保存文件名")#--------------保存文件名标签------------
        self.save_label.grid(row=1, column=0, pady=10, padx=10)
        self.save_label.grid_forget()
        self.notice_label=ttk.Label(self.division_face, text="(如：分账.xlsx)")#------------保存文件提示标签
        self.notice_label.grid(row=2, column=1, padx=10, pady=10)
        self.notice_label.grid_forget()

        #输入框组件
        self.save_entry=ttk.Entry(master=self.division_face, width=35)#--------------保存文件名输入框---------
        self.save_entry.grid(row=1, column=1, padx=10, pady=10)
        self.save_entry.grid_forget()

        #消息框组件
        self.choose_file_text=ttk.Text(self.division_face, height=1, width=35)#--------------选择文件消息框-----------
        self.choose_file_text.grid(row=0, column=1, padx=10, pady=10)
        self.message_box=ttk.Text(self.division_face, height=5, width=35)#--------------结果显示消息框-----------
        self.message_box.grid(row=3, column=1, pady=10, padx=10)
        self.message_box.config(state=DISABLED)
    #-----------控制方法-------------

    def choose_file(self):
        file_path=filedialog.askopenfilename(filetypes=[("需要分账的文件", "*.csv"),["需要分账的文件","*.xlsx"]])
        if file_path: # 如果选择了文件，在选择文件消息框显示选择的文件名,显示保存文件标签和输入框，以及制作收入分账按钮
            self.file_path=file_path
            self.choose_file_text.delete('1.0', 'end')
            self.choose_file_text.insert(END, file_path.split("/")[-1])
            self.save_label.grid(row=1, column=0, pady=10, padx=10)
            self.save_entry.grid(row=1, column=1, padx=10, pady=10)
            self.create_division.grid(row=1, column=2, pady=10, padx=10)
            self.notice_label.grid(row=2, column=1, padx=10, pady=10)

    def divide(self):
        if self.save_entry.get():
            data = []
            with open(self.file_path) as csvfile:
                csv_reader = csv.reader(csvfile)  # 使用csv.reader读取csvfile中的文件
                #header = next(csv_reader)        # 读取第一行每一列的标题
                for row in csv_reader:  # 将csv 文件中的数据保存到data中
                    data.append(row[5])  # 选择某一列加入到data数组中
                print(data)
        else:
            print("no")

    def to_iniface(self):
        self.division_face.destroy()
        init_face.init_face(self.master)