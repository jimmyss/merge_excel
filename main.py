# -*- coding: utf-8 -*-
import tkinter as tk
from faces.init_face import init_face


class basedesk():#主界面
    def __init__(self, master):
        self.root = master
        self.root.config()
        self.root.title('Base page')
        self.root.geometry('750x400')

        init_face(self.root)

if __name__ == '__main__':
    root = tk.Tk()
    basedesk(root)
    root.mainloop()