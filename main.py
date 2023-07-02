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
        self.entry.configure(font=("Helvetica", 14))
        self.vscode_windows = self.get_vscode_windows()
        self.listbox = tk.Listbox(root, font=("Helvetica", 14), width=40)
        self.update_list()
        self.entry.bind('<Return>', self.switch_window)
        self.entry.pack()
        self.listbox.pack()
        self.this_program_window = None

        self.switch_button = tk.Button(root, text="Switch", command=self.switch_window_button)
        self.switch_button.pack()
        id1 = manager.register_hotkey([Key.alt_l, Key.space], None, self.activate)
        manager.suppress = True
        
        self.root.title("My Switcher Program")
        self.entry.focus()

        # 隐藏标题栏
        self.root.overrideredirect(True)

        # 将窗口居中
        window_width = self.root.winfo_reqwidth() + 500
        window_height = self.root.winfo_reqheight()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x_coord = screen_width // 2 - window_width // 2
        y_coord = screen_height // 2 - window_height // 2

        self.root.geometry(f'{window_width}x{window_height}+{x_coord}+{y_coord}')
        

    def activate(self):
        self.update_vscode_windows()  # 更新 vscode 窗口列表
        this_program_window = self.get_this_program_window()
        print(this_program_window)
        this_program_window.minimize()
        this_program_window.restore()
        self.entry.focus()

    def get_vscode_windows(self):
        return [win for win in gw.getAllWindows() if "Visual Studio Code" in win.title]

    def update_vscode_windows(self):
        self.vscode_windows = self.get_vscode_windows()
        for w in self.vscode_windows:
            print(w.title)
        self.update_list()

    def get_this_program_window(self):
        return [win for win in gw.getAllWindows() if "My Switcher Program" in win.title][0]

    def get_acronym_score(self, search, title):
        words = re.findall(r'\b\w', title) 
        acronym = ''.join(words)

        search_idx = 0
        for char in acronym:
            if search_idx < len(search) and char.lower() == search[search_idx].lower():
                search_idx += 1

        return search_idx

    def update_list(self, *args):
        search = self.entry_var.get()
        keywords = search.split()
        score_list = []

        for i, win in enumerate(self.vscode_windows):
            base_score = sum(int(re.search(keyword, win.title, re.I) is not None) for keyword in keywords)
           
            acronym_score = self.get_acronym_score(search, win.title)

            score = base_score + acronym_score
            score_list.append((score, win.title))

        score_list.sort(reverse=True, key=lambda x: x[0])

        self.listbox.delete(0, tk.END)
        for score, title in score_list:
            if score > 0:
                self.listbox.insert(tk.END, title)

        if len(self.listbox.get(0, tk.END)) > 0:
            self.listbox.select_set(0)

    def switch_window(self, event):
        this_program_window = self.get_this_program_window()
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