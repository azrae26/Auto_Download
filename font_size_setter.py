import win32gui
import win32con
from pywinauto.application import Application
import time
from utils import (debug_print, ensure_foreground_window, calculate_center_position, 
                  program_moving_context, check_mouse_before_move)
import pyautogui

@check_mouse_before_move
def set_font_size():
    """設定字型大小"""
    try:
        # 找到目標視窗
        windows = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "stocks" in title.lower():
                    windows.append((hwnd, title))
        win32gui.EnumWindows(callback, windows)
        
        if not windows:
            debug_print("錯誤: 找不到目標視窗")
            return False
            
        # 將視窗帶到前景
        hwnd = windows[0][0]
        if not ensure_foreground_window(hwnd, windows[0][1]):
            debug_print("警告: 無法確保視窗在前景，但將繼續執行")
        
        # 找到字型大小下拉選單
        app = Application(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        combo = main_window.child_window(title="字型大小:", control_type="ComboBox")
        
        debug_print("開始設定字型大小...", color='light_cyan')
        
        # 點擊下拉選單右側按鈕
        rect = combo.rectangle()
        center_x, center_y = calculate_center_position(rect)
        if center_x is None or center_y is None:
            debug_print("無法計算下拉選單位置", color='light_red')
            return False
            
        # 調整到右側按鈕位置
        click_x = rect.right - 10
        click_y = center_y
        
        with program_moving_context():
            # 移動滑鼠到按鈕位置
            pyautogui.moveTo(click_x, click_y, duration=0.2)
            time.sleep(0.5)
            
            # 第一次點擊展開下拉選單
            pyautogui.click()
            time.sleep(1)
            
            # 按向下鍵
            pyautogui.press('down')
            time.sleep(2)
            
            # 按向上鍵
            pyautogui.press('up')
            time.sleep(1)
            
            # 再次點擊收起下拉選單
            pyautogui.click()
        
        debug_print("字型大小設定完成")
        return True
        
    except Exception as e:
        debug_print(f"設定字型大小時發生錯誤: {str(e)}", color='light_red')
        return False