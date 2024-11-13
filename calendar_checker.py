import win32gui
from pywinauto.application import Application
from datetime import datetime
import pyautogui
import time
from utils import debug_print, find_window_handle

class CalendarChecker:
    """日曆檢查器類"""
    HEADER_HEIGHT = 25
    WEEKDAY_HEIGHT = 25
    GRID_ROWS = 6
    GRID_COLS = 7
    CLICK_DELAY = 0.2

    def __init__(self):
        self.hwnd = None
        self.window_title = None
        self.calendar = None
        
    def find_window(self):
        """找到目標視窗"""
        windows = find_window_handle(self.window_title)
        
        if not windows:
            debug_print("錯誤: 找不到目標視窗")
            return False
            
        self.hwnd = windows[0][0]
        self.window_title = windows[0][1]
        debug_print(f"使用視窗: {self.window_title}")
        return True

    def find_calendar(self):
        """找到日曆元素"""
        try:
            app = Application(backend="uia").connect(handle=self.hwnd)
            main_window = app.window(handle=self.hwnd)
            self.calendar = main_window.child_window(
                class_name="WindowsForms10.SysMonthCal32.app.0.32f6d92_r8_ad1"
            )
            return self.calendar.exists()
        except Exception as e:
            debug_print(f"尋找日曆元素時發生錯誤: {str(e)}")
            return False

    def calculate_click_position(self):
        """計算點擊位置"""
        rect = self.calendar.rectangle()
        today = datetime.now()
        
        # 計算網格尺寸
        cell_width = (rect.right - rect.left) / self.GRID_COLS
        cell_height = (rect.bottom - rect.top - self.HEADER_HEIGHT - self.WEEKDAY_HEIGHT) / self.GRID_ROWS
        
        # 計算日期位置
        first_day = today.replace(day=1)
        first_day_col = (first_day.weekday() + 1) % 7
        days_from_start = today.day - 1
        total_position = first_day_col + days_from_start
        
        # 計算目標座標
        target_row = total_position // 7
        target_col = total_position % 7
        
        x = rect.left + (target_col + 0.5) * cell_width
        y = rect.top + self.HEADER_HEIGHT + self.WEEKDAY_HEIGHT + (target_row + 0.5) * cell_height
        
        return int(x), int(y)

    def click_today(self):
        """點擊今日日期"""
        try:
            x, y = self.calculate_click_position()
            debug_print(f"今天是 {datetime.now().day} 號")
            debug_print(f"計算得出的點擊位置: x={x}, y={y}")
            
            # 執行三次點擊
            for _ in range(3):
                pyautogui.click(x, y)
                time.sleep(self.CLICK_DELAY)
            
            debug_print("已執行三次點擊")
            return True
            
        except Exception as e:
            debug_print(f"點擊日期時發生錯誤: {str(e)}")
            return False

def start_calendar_checker():
    """主函數"""
    debug_print("開始檢測日歷元素並點選今日日期...")
    
    checker = CalendarChecker()
    if not checker.find_window():
        return
        
    try:
        win32gui.SetForegroundWindow(checker.hwnd)
        time.sleep(0.2)
    except Exception as e:
        debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")
    
    if not checker.find_calendar():
        debug_print("無法找到日歷元素")
        return
        
    debug_print("找到日歷元素!")
    if checker.click_today():
        try:
            debug_print(f"日歷可見性: {checker.calendar.is_visible()}")
        except Exception as e:
            debug_print(f"獲取日歷資訊時發生錯誤: {str(e)}")