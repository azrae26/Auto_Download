import keyboard
import pywinauto
from pywinauto.application import Application
import win32gui
import threading
from datetime import datetime
import pyautogui
import time

def debug_print(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f"[{timestamp}] {message}"
    print(formatted_message)

def find_window_handle(target_title=None):
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if target_title:
                if target_title.lower() in title.lower():
                    windows.append((hwnd, title))
            else:
                if title:
                    windows.append((hwnd, title))
    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows

def check_calendar():
    try:
        # 尋找目標視窗
        target_windows = find_window_handle("stocks")
        if not target_windows:
            debug_print("錯誤: 找不到相關視窗")
            return

        hwnd, window_title = target_windows[0]
        debug_print(f"使用視窗: {window_title}")

        # 將視窗帶到前景
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.2)
        except Exception as e:
            debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")

        # 連接到應用程式並找到日曆元素
        app = Application(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        calendar = main_window.child_window(class_name="WindowsForms10.SysMonthCal32.app.0.32f6d92_r8_ad1")
        
        if calendar.exists():
            debug_print("找到日歷元素!")
            
            try:
                # 獲取日歷的位置
                rect = calendar.rectangle()
                debug_print(f"日歷位置: {rect}")
                
                # 直接使用原始座標
                calendar_left = rect.left
                calendar_top = rect.top
                calendar_width = rect.right - rect.left
                calendar_height = rect.bottom - rect.top
                
                # 獲取今天的日期和月份資訊
                today = datetime.now()
                current_day = today.day
                first_day = today.replace(day=1)
                
                debug_print(f"今天是 {current_day} 號")
                
                # 計算日期網格參數
                HEADER_HEIGHT = 25
                WEEKDAY_HEIGHT = 25
                ROWS = 6
                COLS = 7
                
                # 計算單個日期格子的大小
                cell_width = calendar_width / COLS
                cell_height = (calendar_height - HEADER_HEIGHT - WEEKDAY_HEIGHT) / ROWS
                
                # 計算第一天的位置
                first_day_weekday = first_day.weekday()
                first_day_col = (first_day_weekday + 1) % 7
                
                # 計算目標日期的位置
                days_from_start = current_day - 1
                total_position = first_day_col + days_from_start
                target_row = total_position // 7
                target_col = total_position % 7
                
                # 計算點擊座標
                x = calendar_left + (target_col + 0.5) * cell_width
                y = calendar_top + HEADER_HEIGHT + WEEKDAY_HEIGHT + (target_row + 0.5) * cell_height
                
                debug_print(f"計算得出的點擊位置: x={int(x)}, y={int(y)}")
                
                # 執行雙擊
                pyautogui.click(int(x), int(y))
                time.sleep(0.05)
                pyautogui.click(int(x), int(y))
                
                debug_print("已執行雙擊")
                
            except Exception as e:
                debug_print(f"處理日歷時發生錯誤: {str(e)}")
            
            return calendar
        
        debug_print("無法找到日歷元素")
        
    except Exception as e:
        debug_print(f"檢查日歷時發生錯誤: {str(e)}")

def start_calendar_checker():
    debug_print("開始檢測日歷元素並點選今日日期...")
    calendar = check_calendar()
    if calendar:
        try:
            debug_print(f"日歷可見性: {calendar.is_visible()}")
        except Exception as e:
            debug_print(f"獲取日歷資訊時發生錯誤: {str(e)}")