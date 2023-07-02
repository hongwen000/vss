from PyHotKey import Key, keyboard_manager as manager

import pygetwindow as gw
import tkinter as tk
from pynput import keyboard
import re
import ctypes
import pyautogui


class App:

    def __init__(self, root):
        self.root = root
        self.entry_var = tk.StringVar()
        self.entry_var.trace("w", self.update_list)
        self.entry = tk.Entry(root, textvariable=self.entry_var)
        self.vscode_windows = self.get_vscode_windows()
        self.listbox = tk.Listbox(root)
        self.update_list()
        self.entry.bind('<Return>', self.switch_window)
        self.entry.pack()
        self.listbox.pack()
        self.this_program_window = None


        self.switch_button = tk.Button(root, text="Switch", command=self.switch_window_button)
        self.switch_button.pack()
        # self.root.attributes('-topmost', True)  # Add this line
        id1 = manager.register_hotkey([Key.alt_l, Key.space], None, self.activate)
        manager.suppress = True
        self.root.title("My Switcher Program")  # 设置程序窗口标题
        self.entry.focus()
        # 获取并存储自身的窗口




    def activate(self):
        # 激活程序窗口
        this_program_window = [win for win in gw.getAllWindows() if "My Switcher Program" in win.title][0]
        print(this_program_window)
        this_program_window.minimize()
        this_program_window.restore()
        # 设置焦点在输入框上
        self.entry.focus()



    def get_vscode_windows(self):
        return [win for win in gw.getAllWindows() if "Visual Studio Code" in win.title]

    def update_list(self, *args):
        search = self.entry_var.get()
        self.listbox.delete(0, tk.END)
        for i, win in enumerate(self.vscode_windows):
            if re.search(search, win.title, re.I):
                self.listbox.insert(tk.END, win.title)
        if len(self.listbox.get(0, tk.END)) > 0:
            self.listbox.select_set(0)

    def switch_window(self, event):
        this_program_window = [win for win in gw.getAllWindows() if "My Switcher Program" in win.title][0]
        this_program_window.minimize()
        selection = self.listbox.curselection()
        if selection:
            title = self.listbox.get(selection[0])
            for win in self.vscode_windows:
                if win.title == title:
                    win.restore()
                    win.activate()
                    break
        self.entry.focus()

    def switch_window_button(self):
        self.switch_window(None)

def main():
    root = tk.Tk()
    app = App(root)

    root.mainloop()

from ctypes import windll
if __name__ == "__main__":
    windll.shcore.SetProcessDpiAwareness(1)
    main()
