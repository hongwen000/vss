import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def create_executable(script):
    subprocess.check_call([sys.executable, "-m", "PyInstaller", "--onefile", "--windowed", script])

if __name__ == "__main__":
    # 安装需要的包
    install("PyInstaller")
    install("PyHotKey")
    install("pygetwindow")
    install("pynput")
    install("pyautogui")

    # 创建可执行文件
    create_executable("main.py")
