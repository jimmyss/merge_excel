import os
import sys
import tkinter as tk
def get_path():
    app_path = ""
    if getattr(sys, 'frozen', False):
        app_path = os.path.dirname(sys.executable).replace("\\", "/")
    elif __file__:
        app_path = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")
    return app_path

class re_Text():

    def __init__(self, queue):
        self.queue = queue

    def write(self, content):
        self.queue.put(content)

class GressBar():

	def start(self):
		top = tk.Toplevel()
		self.master = top
		top.overrideredirect(True)
		top.title("进度条")
		tk.Label(top, text="任务正在运行中,请稍等……", fg="green").pack(pady=2)
		prog = tk.ttk.Progressbar(top, mode='indeterminate', length=200)
		prog.pack(pady=10, padx=35)
		prog.start()

		top.resizable(False, False)
		top.update()
		curWidth = top.winfo_width()
		curHeight = top.winfo_height()
		scnWidth, scnHeight = top.maxsize()
		tmpcnf = '+%d+%d' % ((scnWidth - curWidth) / 2, (scnHeight - curHeight) / 2)
		top.geometry(tmpcnf)
		top.mainloop()

	def quit(self):
		if self.master:
			self.master.destroy()



class Entry_(tk.Entry):
    def __init__(self, master, placeholder, **kw):
        super().__init__(master, **kw)

        self.placeholder = placeholder
        self._is_password = True if placeholder == "password" else False

        self.bind("<FocusIn>", self.on_focus_in)
        self.bind("<FocusOut>", self.on_focus_out)

        self._state = 'placeholder'
        self.insert(0, self.placeholder)

    def on_focus_in(self, event):
        if self._is_password:
            self.configure(show='*')

        if self._state == 'placeholder':
            self._state = ''
            self.delete('0', 'end')

    def on_focus_out(self, event):
        if not self.get():
            if self._is_password:
                self.configure(show='')

            self._state = 'placeholder'
            self.insert(0, self.placeholder)