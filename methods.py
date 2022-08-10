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