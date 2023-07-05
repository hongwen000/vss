# import subprocess
# import sys

# def install(package):
#     subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# def create_executable(script):
#     subprocess.check_call([sys.executable, "-m", "PyInstaller", "--onefile", "--windowed", script])

# if __name__ == "__main__":
#     # 安装需要的包
#     install("PyInstaller")
#     install("PyHotKey")
#     install("pygetwindow")
#     install("pynput")
#     install("pyautogui")

#     # 创建可执行文件
#     create_executable("main.py")

import os
import sys
import ctypes
from ctypes import wintypes, windll
import wpath

from PIL import Image, ImageDraw

image = Image.new('RGB', (64, 64), color = 'white')

# 在图片上绘制一个黑色的矩形
draw = ImageDraw.Draw(image)
draw.rectangle(
    [(10, 10), (54, 54)],
    fill='black'
)

# 保存为.ico文件
image.save('icon.ico')

import os
import sys
import winshell
from pathlib import Path
from wpath import FOLDERID
import wpath

# 获取Python解释器的路径
python_path = sys.executable

# 获取pythonw.exe的路径
pythonw_path = os.path.join(os.path.dirname(python_path), "pythonw.exe")

# 获取main.py的路径
main_path = os.path.join(os.getcwd(), "main.py")

# 获取桌面的路径
desktop = Path(wpath.get_path(FOLDERID.Desktop))

# 设置图标文件的路径
icon = os.path.join(os.getcwd(), "icon.ico")

# 创建快捷方式的命令
shortcut_cmd = f'{pythonw_path} {main_path}'

# 创建桌面快捷方式
desktop_shortcut_path = str(desktop / "MySwitcherProgram.lnk")
with winshell.shortcut(desktop_shortcut_path) as link:
    link.path = pythonw_path
    link.arguments = main_path
    link.icon_location = (icon, 0)
    link.description = "My Switcher Program"

# 获取开始菜单的路径
start_menu = Path(wpath.get_path(FOLDERID.StartMenu))

# 创建开始菜单快捷方式
start_menu_shortcut_path = str(start_menu / "MySwitcherProgram.lnk")
with winshell.shortcut(start_menu_shortcut_path) as link:
    link.path = pythonw_path
    link.arguments = main_path
    link.icon_location = (icon, 0)
    link.description = "My Switcher Program"
