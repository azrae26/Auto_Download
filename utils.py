from datetime import datetime
from queue import Queue
import win32gui

debug_queue = Queue()

def debug_print(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f"[{timestamp}] {message}"
    debug_queue.put(formatted_message)
    print(formatted_message)

def find_window_handle(target_title=None):
    """
    找到視窗，使用模糊比對
    target_title: 目標視窗標題（不含版本號）
    """
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # 如果有指定目標標題，使用模糊比對
            if target_title:
                # 移除版本號部分再比對
                window_name = title.split('版本')[0].strip()
                if target_title.lower() in window_name.lower():
                    windows.append((hwnd, title))
            else:
                if title:
                    windows.append((hwnd, title))
    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows 