from PyHotKey import Key, keyboard_manager as manager
import json
import pygetwindow as gw
import tkinter as tk
from pynput import keyboard
import re,os
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

def Get_HWND_DPI(window_handle):
    # To detect high DPI displays and avoid need to set Windows compatibility flags
    import os

    if os.name == "nt":
        from ctypes import windll, pointer, wintypes

        try:
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass  # this will fail on Windows Server and maybe early Windows
        DPI100pc = 96  # DPI 96 is 100% scaling
        DPI_type = 0  # MDT_EFFECTIVE_DPI = 0, MDT_ANGULAR_DPI = 1, MDT_RAW_DPI = 2
        winH = wintypes.HWND(window_handle)
        monitorhandle = windll.user32.MonitorFromWindow(
            winH, wintypes.DWORD(2)
        )  # MONITOR_DEFAULTTONEAREST = 2
        X = wintypes.UINT()
        Y = wintypes.UINT()
        try:
            windll.shcore.GetDpiForMonitor(
                monitorhandle, DPI_type, pointer(X), pointer(Y)
            )
            return X.value, Y.value, (X.value + Y.value) / (2 * DPI100pc)
        except Exception:
            return 96, 96, 1  # Assume standard Windows DPI & scaling
    else:
        return None, None, 1  # What to do for other OSs?


def TkGeometryScale(s, cvtfunc):
    patt = r"(?P<W>\d+)x(?P<H>\d+)\+(?P<X>\d+)\+(?P<Y>\d+)"  # format "WxH+X+Y"
    R = re.compile(patt).search(s)
    G = str(cvtfunc(R.group("W"))) + "x"
    G += str(cvtfunc(R.group("H"))) + "+"
    G += str(cvtfunc(R.group("X"))) + "+"
    G += str(cvtfunc(R.group("Y")))
    return G


def MakeTkDPIAware(TKGUI):
    TKGUI.DPI_X, TKGUI.DPI_Y, TKGUI.DPI_scaling = Get_HWND_DPI(TKGUI.winfo_id())
    logger.debug("DPI scaling: %s" % TKGUI.DPI_scaling)
    TKGUI.TkScale = lambda v: int(float(v) * TKGUI.DPI_scaling)
    TKGUI.TkGeometryScale = lambda s: TkGeometryScale(s, TKGUI.TkScale)

import logging
import sys
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(logging.StreamHandler(sys.stdout))

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

        self.custom_terms = {"r743", "ubuntu", "msm", "aosp_host_working_dir", "aosp", "apk", "androidtools", "vss"}  # Add your custom terms here

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
        # with open('log.txt', 'a+') as f:
        #     logger.debug("click_select", self.listbox.curselection(), self.ignore, file=f)
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
        MakeTkDPIAware(self.root)
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
        #     logger.debug(w.title)
        self.update_list()

    def get_this_program_window(self):
        return [win for win in gw.getAllWindows() if "My Switcher Program" in win.title][0]


    def nj_split(self, word):
        if word in self.custom_terms:
            return [word]
        else:
            return wordninja.split(word)

    def split_title(self, title):
        logger.debug(title)
        # Split the title by " - " and process each part
        parts = title.split(" - ")
        logger.debug(parts)

        # Process each part
        processed_parts = []
        part: str
        for part in parts:
            no_term = False
            # Iteratively check for custom terms and split them out
            while not no_term:
                term_idxs = [part.find(term) for term in self.custom_terms]
                no_term = all(idx == -1 for idx in term_idxs)
                if not no_term:
                    first_term = min(list(filter(lambda x: x >= 0, term_idxs)))
                    length = 0
                    target_term = None
                    for j, __term in enumerate(self.custom_terms):
                        if term_idxs[j] == first_term:
                            if len(__term) > length:
                                length = len(__term)
                                target_term = __term
                    if target_term:
                        before, term, after = part.partition(target_term)
                        if before:
                            sp = self.nj_split(before)
                            processed_parts.extend(sp)  # Use wordninja to split
                            logger.debug("before: %s, after: %s", before, sp)
                        processed_parts.append(term)  # Add the custom term
                        logger.debug(processed_parts)
                        part = after  # Continue processing the rest of the string

            # Add remaining part if any, split using wordninja
            if part:
                processed_parts.extend(self.nj_split(part))
                logger.debug(processed_parts)

        # Filter out empty strings and return the result
        return list(filter(lambda x: len(x) > 0, processed_parts))


    def get_acronym_score(self, search, title):
        words = self.split_title(title)
        title = " ".join(words)
        words = list(filter(lambda x: len(x) > 0, title.split(" ")))
        map_acronym_to_word = [{word[0]: word} for word in words]
        acronym = [word[0] for word in words]

        # Start searching
        logger.debug("title: %s", title)
        matched_length = self.search_acronym(search, acronym, map_acronym_to_word)
        return matched_length  # Increase the score by matched length squared


    def search_acronym(self, search, acronym, map_acronym_to_word, search_idx=0, acronym_idx=0):
        # Recursion base case: if all search chars are found, return len(search)
        if search_idx == len(search):
            return len(search)
        
        # Recursion base case: if all acronyms are used and search not found, return 0
        if acronym_idx == len(acronym):
            return 0

        bonus = 0
        logger.debug("search_acronym with %s[%s] %s against %s[%s] %s", search, search_idx, search[search_idx], acronym, acronym_idx, acronym[acronym_idx])
        max_match_len = 0
        # If this acronym char match this search char
        if search[search_idx].lower() == acronym[acronym_idx].lower():
            # Try to match next search char in this word first (greedy)
            word = map_acronym_to_word[acronym_idx][acronym[acronym_idx]]
            word_match_len, whole_match = self.search_in_word(search, search_idx, word)
            bonus = 1 if whole_match else 0
            logger.debug("matched %s with %s in %s, word_match_len=%s", search[search_idx], acronym[acronym_idx], word, word_match_len)

            # Try to match rest of search string with rest of acronyms
            for end_idx in range(search_idx + word_match_len, len(search) + 1):
                next_match_len = self.search_acronym(search, acronym, map_acronym_to_word, end_idx, acronym_idx + 1)
                max_match_len = max(max_match_len, (end_idx - search_idx) + next_match_len)

        # Also try without matching this search char and move to next acronym
        max_match_len = max(max_match_len, self.search_acronym(search, acronym, map_acronym_to_word, search_idx, acronym_idx + 1))

        return max_match_len + bonus

    def search_in_word(self, search, search_idx, word):
        # Search for search[search_idx:] in word, return the length of matched substring
        word_idx = 0
        matched_len = 0
        word_match = True
        while search_idx < len(search) and word_idx < len(word):
            if search[search_idx].lower() == word[word_idx].lower():
                matched_len += 1
                search_idx += 1
            else:
                word_match = False
            word_idx += 1
        if ((search_idx < len(search) and search[search_idx] == ' ') or (search_idx == len(search))) and (word_match):
            return matched_len, True
        else:
            return matched_len, False


    def get_acronym_score_old(self, search, title):
        # logger.debug(title)
        words = [" ".join(wordninja.split(part)) for part in title.split(" - ")]
        # logger.debug(words)
        title = " ".join(words)
        word_starts = re.findall(r'\b\w', title) 
        # logger.debug(word_starts)
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

    def load_magic_searches(self):
        # get script directory
        d = os.path.dirname(os.path.realpath(__file__))
        with open(f'{d}/config.json', 'r') as file:
            data = json.load(file)
        return data

    def update_list(self, *args):

        magic_searches = self.load_magic_searches()
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
            
                acronym_score = self.get_acronym_score(search, search_title) * 10

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
                    logger.debug("%s: %s", title, score)

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
                        logger.debug(app_spec)
                        win_spec = app_spec.window(handle=win._hWnd)
                        logger.debug(win_spec)
                        win_spec.set_focus()
                    except Exception as e:
                        logger.debug("Exception")
                        logger.debug(e)
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