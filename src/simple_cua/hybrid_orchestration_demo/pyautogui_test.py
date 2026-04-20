# pyautogui_test.py

import time
import pyautogui
import os
import subprocess
from PIL import ImageGrab
import pygetwindow as gw

# Configuration
data_path = r"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data"
capture_window_only = True  # Set to False to capture full screen

os.makedirs(data_path, exist_ok=True)

# Open a new CMD window
subprocess.Popen('start cmd.exe', shell=True)
print("New CMD window opened. Waiting...")
time.sleep(2)

print("Focus the CMD window now...")
time.sleep(2)

# Type the command
pyautogui.write(f"type {os.path.join(data_path, 'invoice.txt')}\n", interval=0.05)
time.sleep(1)

# Capture screenshot
try:
    if capture_window_only:
        cmd_windows = gw.getWindowsWithTitle('cmd.exe')
        if cmd_windows:
            cmd_window = cmd_windows[0]
            bbox = (cmd_window.left, cmd_window.top, cmd_window.right, cmd_window.bottom)
            screenshot = ImageGrab.grab(bbox=bbox)
        else:
            print("CMD window not found, capturing full screen")
            screenshot = ImageGrab.grab()
    else:
        screenshot = ImageGrab.grab()
except Exception as e:
    print(f"Error: {e}, capturing full screen instead")
    screenshot = ImageGrab.grab()

screenshot.save(os.path.join(data_path, "terminal_test.png"))
print("Screenshot saved as terminal_test.png")

