from PyHotKey import Key, keyboard_manager as manager
import pygetwindow as gw
import tkinter as tk
from pynput import keyboard
import re
import ctypes
import pyautogui
from pywinauto import Desktop, Application
from PIL import Image, ImageDraw
from pprint import pprint
import win32con
import wordninja
import win32api
import ctypes
from win32api import (
    GetCurrentThreadId
)
from ctypes import *
import ctypes
import ctypes.wintypes
from ctypes import (
    POINTER,
    Structure
)
from ctypes.wintypes import (
    DWORD,
    LONG,
    WORD,
    BYTE,
    RECT,
    UINT,
    ATOM
)
class tagTITLEBARINFO(Structure):
    pass
tagTITLEBARINFO._fields_ = [
    ('cbSize', DWORD),
    ('rcTitleBar', RECT),
    ('rgstate', DWORD * 6),
]
PTITLEBARINFO = POINTER(tagTITLEBARINFO)
LPTITLEBARINFO = POINTER(tagTITLEBARINFO)
TITLEBARINFO = tagTITLEBARINFO
only_show_vscode = True
class tagWINDOWINFO(Structure):
    def __str__(self) -> str:
        return '\n'.join([key + ':' + str(getattr(self, key)) for key, value in self._fields_])

tagWINDOWINFO._fields_ = [
    ('cbSize', DWORD),
    ('rcWindow', RECT),
    ('rcClient', RECT),
    ('dwStyle', DWORD),
    ('dwExStyle', DWORD),
    ('dwWindowStatus', DWORD),
    ('cxWindowBorders', UINT),
    ('cyWindowBorders', UINT),
    ('atomWindowType', ATOM),
    ('wCreatorVersion', WORD),
]
WINDOWINFO = tagWINDOWINFO
LPWINDOWINFO = POINTER(tagWINDOWINFO)
PWINDOWINFO = POINTER(tagWINDOWINFO)

def is_alt_tab_window(hwnd: int) -> bool:
    """Check whether a window is shown in alt-tab.

    See http://stackoverflow.com/a/7292674/238472 for details.
    """
    hwnd_walk = win32con.NULL
    hwnd_try = ctypes.windll.user32.GetAncestor(hwnd, win32con.GA_ROOTOWNER)
    while hwnd_try != hwnd_walk:
        hwnd_walk = hwnd_try
        hwnd_try = ctypes.windll.user32.GetLastActivePopup(hwnd_walk)
        if gw.Win32Window(hwnd_try).visible:
            break

    if hwnd_walk != hwnd:
        return False

    # the following removes some task tray programs and "Program Manager"
    ti = TITLEBARINFO()
    ti.cbSize = ctypes.sizeof(ti)
    ctypes.windll.user32.GetTitleBarInfo(hwnd, ctypes.byref(ti))
    if ti.rgstate[0] & win32con.STATE_SYSTEM_INVISIBLE:
        return False

    # Tool windows should not be displayed either, these do not appear in the
    # task bar.
    if win32api.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOOLWINDOW:
        return False

    pwi = WINDOWINFO()
    windll.user32.GetWindowInfo(hwnd, byref(pwi))
    # A top-level window created with this style does not become the foreground
    # window when the user clicks it. The system does not bring this window to
    # the foreground when the user minimizes or closes the foreground window.
    # The window does not appear on the taskbar by default.
    if pwi.dwExStyle & win32con.WS_EX_NOACTIVATE:
        return False

    return True

class App:
    def __init__(self, root, tray_icon) -> None:
        self.recent: str = ""
        self.root = root
        self.tray_icon = tray_icon
        self.entry_var = tk.StringVar()
        self.entry_var.trace("w", self.update_list)
        self.entry = tk.Entry(root, textvariable=self.entry_var)
        self.entry.configure(font=("Helvetica", 14))
        self.window_list = self.get_windows()
        self.listbox = tk.Listbox(root, font=("Helvetica", 14), width=80)
        self.update_list()

        self.entry.bind('<Return>', self.switch_window)
        self.entry.bind('<Escape>', self.cancel_window)
        self.entry.bind('<KeyRelease-Up>', self.select_up)
        self.entry.bind('<KeyRelease-Down>', self.select_down)

        self.listbox.bind('<Up>', self.select_up)
        self.listbox.bind('<Down>', self.select_down)
        self.listbox.bind('<Return>', self.switch_window)
        self.listbox.bind('<Escape>', self.cancel_window)
        # self.listbox.bind('<<ListboxSelect>>', self.click_select)

        self.entry.pack()
        self.listbox.pack()
        self.this_program_window = None

        self.switch_button = tk.Button(root, text="Quit", command=self.exit_tkinter)
        self.switch_button.pack()
        id1 = manager.register_hotkey([Key.alt_l, Key.space], None, self.activate)
        manager.suppress = True
        
        self.update_window_list()  # 更新 vscode 窗口列表
        self.root.title("My Switcher Program")
        self.entry.focus()

        # 隐藏标题栏
        self.root.overrideredirect(True)

        # 将窗口居中
        window_width = self.root.winfo_reqwidth() + 1200
        window_height = self.root.winfo_reqheight() + 200
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x_coord = screen_width // 2 - window_width // 2
        y_coord = screen_height // 2 - window_height // 2

        self.root.geometry(f'{window_width}x{window_height}+{x_coord}+{y_coord}')
        self.ignore = True

    def click_select(self, event):
        with open('log.txt', 'a+') as f:
            print("click_select", self.listbox.curselection(), self.ignore, file=f)
        if not self.ignore:
            self.ignore = True
            self.switch_window(event)

    def select_up(self, event):
        self.ignore = True
        current_selection = self.listbox.curselection()
        if current_selection:
            self.listbox.select_clear(current_selection)
            self.listbox.select_set(current_selection[0]-1)
            self.listbox.see(current_selection[0]-1) # Ensure new selection is visible

    def select_down(self, event):
        current_selection = self.listbox.curselection()
        if current_selection:
            self.listbox.select_clear(current_selection)
            self.listbox.select_set(current_selection[0]+1)
            self.listbox.see(current_selection[0]+1) # Ensure new selection is visible
        

    def activate(self):
        self.update_window_list()  # 更新 vscode 窗口列表
        this_program_window = self.get_this_program_window()
        this_program_window.minimize()
        this_program_window.restore()
        self.entry.focus()

    def get_windows(self):
        win: gw.Win32Window
        return [win for win in gw.getAllWindows() if (len(win.title) > 0 and is_alt_tab_window(win._hWnd))]

    def update_window_list(self):
        self.window_list = self.get_windows()
        # for w in self.window_list:
        #     print(w.title)
        self.update_list()

    def get_this_program_window(self):
        return [win for win in gw.getAllWindows() if "My Switcher Program" in win.title][0]

    def get_acronym_score(self, search, title):
        words = [" ".join(wordninja.split(part)) for part in title.split(" - ")]
        title = " ".join(words)
        words = list(filter(lambda x: len(x) > 0, title.split(" ")))
        map_acronym_to_word = [{word[0]: word} for word in words]
        acronym = [word[0] for word in words]

        # Start searching
        matched_length = self.search_acronym(search, acronym, map_acronym_to_word)
        return matched_length  # Increase the score by matched length squared


    def search_acronym(self, search, acronym, map_acronym_to_word, search_idx=0, acronym_idx=0):
        # Recursion base case: if all search chars are found, return len(search)
        if search_idx == len(search):
            return len(search)
        
        # Recursion base case: if all acronyms are used and search not found, return 0
        if acronym_idx == len(acronym):
            return 0

        max_match_len = 0
        # If this acronym char match this search char
        if search[search_idx].lower() == acronym[acronym_idx].lower():
            # Try to match next search char in this word first (greedy)
            word = map_acronym_to_word[acronym_idx][acronym[acronym_idx]]
            word_match_len = self.search_in_word(search, search_idx, word)

            # Try to match rest of search string with rest of acronyms
            for end_idx in range(search_idx + 1, search_idx + word_match_len + 1):
                next_match_len = self.search_acronym(search, acronym, map_acronym_to_word, end_idx, acronym_idx + 1)
                max_match_len = max(max_match_len, (end_idx - search_idx) + next_match_len)

        # Also try without matching this search char and move to next acronym
        max_match_len = max(max_match_len, self.search_acronym(search, acronym, map_acronym_to_word, search_idx, acronym_idx + 1))

        return max_match_len

    def search_in_word(self, search, search_idx, word):
        # Search for search[search_idx:] in word, return the length of matched substring
        word_idx = 0
        matched_len = 0
        while search_idx < len(search) and word_idx < len(word):
            if search[search_idx].lower() == word[word_idx].lower():
                matched_len += 1
                search_idx += 1
            word_idx += 1
        return matched_len


    def get_acronym_score_old(self, search, title):
        # print(title)
        words = [" ".join(wordninja.split(part)) for part in title.split(" - ")]
        # print(words)
        title = " ".join(words)
        word_starts = re.findall(r'\b\w', title) 
        # print(word_starts)
        acronym = ''.join(word_starts)

        search_idx = 0
        for idx, char in enumerate(acronym):
            if search_idx >= len(search):
                break

            if char.lower() == search[search_idx].lower():
                search_idx += 1

        return search_idx if search_idx == len(search) else 0

    def remove_vscode_postfix(self, title: str) -> str:
        parts = title.split(" - ")
        if len(parts) > 3:
            # Remove the last two parts (profile name and "Visual Studio Code")
            parts = parts[:-2]
        else:
            # Remove the last part ("Visual Studio Code")
            parts = parts[:-1]
        return " - ".join(parts)


    def update_list(self, *args):
        magic_searches = {
            "a": "aosp_r743",
            "s": "kernel_sunfish_r743",
            "g": "kernel_goldfish_r743",
            "t": "androidtools_r743",
            "ra": "aosp_r743",
            "rs": "kernel_sunfish_r743",
            "rg": "kernel_goldfish_r743",
            "rt": "androidtools_r743",
            "ua": "aosp_ubuntu",
            "us": "kernel_sunfish_ubuntu",
            "ug": "kernel_goldfish_ubuntu",
            "ut": "androidtools_ubuntu",
            "ya": "aosp_yzy_r743",
            "ys": "kernel_sunfish_yzy_r743",
            "yg": "kernel_goldfish_yzy_r743",
            "yt": "androidtools_yzy_r743",
        }
        search = self.entry_var.get()
        if search in magic_searches:
            search = magic_searches[search]
        not_vscode, is_vscode = [], []
        is_vscode = [win for win in self.window_list if "Visual Studio Code" in win.title]
        not_vscode = [win for win in self.window_list if win not in is_vscode]
        if only_show_vscode:
            not_vscode.clear()
        if len(search) == 0:
            self.listbox.delete(0, tk.END)
            win_titles = [win.title for win in (is_vscode + not_vscode) if len(win.title) > 0]
            if self.recent in win_titles:
                self.listbox.insert(tk.END, self.recent)
            for win_title in win_titles:
                if win_title != self.recent:
                    self.listbox.insert(tk.END, win_title)
        else:
            keywords = search.split()
            score_list = []
            

            for i, win in enumerate(is_vscode):
                rcu_score = 0
                search_title = self.remove_vscode_postfix(win.title)
                orig_title = win.title
                if orig_title == self.recent:
                    rcu_score = 10
                base_score = sum(int(re.search(keyword, search_title, re.I) is not None) for keyword in keywords)
                base_score *= 100
            
                acronym_score = self.get_acronym_score(search, search_title)

                score = base_score + acronym_score + rcu_score
                score_list.append((score, orig_title))

            for i, win in enumerate(not_vscode):
                rcu_score = 0
                search_title = win.title
                orig_title = win.title
                if orig_title == self.recent:
                    rcu_score = 10
                base_score = sum(int(re.search(keyword, search_title, re.I) is not None) for keyword in keywords)
                base_score *= 100
                score = base_score + 0 + rcu_score
                score_list.append((score, orig_title))

            score_list.sort(reverse=True, key=lambda x: x[0])

            self.listbox.delete(0, tk.END)
            for score, title in score_list:
                if score > 0:
                    self.listbox.insert(tk.END, title)
                    print(f"{title}: {score}")

        if len(self.listbox.get(0, tk.END)) > 0:
            self.listbox.select_set(0)

    def __switch_window(self, selection):
        this_program_window = self.get_this_program_window()
        this_program_window.minimize()

        if selection:
            title = self.listbox.get(selection[0])
            for win in self.window_list:
                if win.title == title:
                    # self.entry.delete(0, tk.END)
                    # self.entry.focus()
                    # left, top, right, bottom = win._getWindowRect()
                    # win.restore()
                    # win.maximize()
                    # result = ctypes.windll.user32.SetWindowPos(win._hWnd, 0, 0, 0, 0, 0, 0x43)
                    # win.activate()

                    try:
                        app_spec = Application(backend='win32').connect(handle=win._hWnd)
                        print(app_spec)
                        win_spec = app_spec.window(handle=win._hWnd)
                        print(win_spec)
                        win_spec.set_focus()
                    except Exception as e:
                        print("Exception")
                        print(e)
                    self.recent = win.title

                    break
        self.entry.delete(0, tk.END)
        self.entry.focus()

    def cancel_window(self, event):
        self.__switch_window(None)

    def switch_window(self, event):
        selection = self.listbox.curselection()
        self.__switch_window(selection)

    def switch_window_button(self):
        self.switch_window(None)

    def exit_tkinter(self):
        self.root.quit()
        self.root.destroy()
        self.tray_icon.stop()

import pystray
import threading
import time, sys

def main():
    quit_event = threading.Event()
    # 创建一个新的白色图片
    image = Image.new('RGB', (64, 64), color = 'white')

    # 在图片上绘制一个黑色的矩形
    draw = ImageDraw.Draw(image)
    draw.rectangle(
        [(10, 10), (54, 54)],
        fill='black'
    )
    def run_icon():
        icon.run()

    menu = ()
    # 现在你可以使用这个图片作为你的系统托盘图标
    icon = pystray.Icon("name", image, "My System Tray Icon", menu)
    icon_thread = threading.Thread(target=run_icon)
    icon_thread.start()

    root = tk.Tk()
    app = App(root, icon)
    root.mainloop()


from ctypes import windll
if __name__ == "__main__":
    windll.shcore.SetProcessDpiAwareness(1)
    main()