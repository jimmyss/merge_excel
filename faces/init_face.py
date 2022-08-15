import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from faces import data_validation, merge_excel, revenue_division, bill_summary

class init_face():
    def __init__(self, master):
        """

        :param master: 主界面窗口
        """
        self.master = master
        self.master.config()
        # 基准界面initface
        self.initface = ttk.Frame(self.master, )
        self.initface.grid(row=0, column=0)

        #初始化基准界面组件
        task1_but = ttk.Button(self.initface, text='excel合并功能', command=self.to_merge_excel)
        task1_but.grid(row=1, column=1, sticky=E, pady=10, padx=10)
        task2_but = ttk.Button(self.initface, text='数据有效性添加', command=self.to_data_validation)
        task2_but.grid(row=2, column=1, sticky=E, pady=10, padx=10)
        task3_but = ttk.Button(self.initface, text='月度分账收入', command=self.to_revenue_division)
        task3_but.grid(row=3, column=1, sticky=E, padx=10, pady=10)
        task4_but = ttk.Button(self.initface, text='月度账单汇总取值', command=self.to_bill_summary)
        task4_but.grid(row=4, column=1, sticky=E, padx=10, pady=10)

    def to_merge_excel(self, ):
        self.initface.destroy()
        merge_excel.merge_excel(self.master)

    def to_data_validation(self):
        self.initface.destroy()
        data_validation.data_validation(self.master)

    def to_revenue_division(self):
        self.initface.destroy()
        revenue_division.revenue_division(self.master)

    def to_bill_summary(self):
        self.initface.destroy()
        bill_summary.bill_summary(self.master)