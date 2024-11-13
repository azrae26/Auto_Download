import win32gui
from pywinauto.application import Application
from datetime import datetime
import pyautogui
import time
from utils import (debug_print, find_window_handle, ensure_foreground_window, 
                  calculate_center_position)

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

    def calculate_click_position(self, days_ago): # self意思是CalendarChecker類的實例
        """計算點擊位置"""
        rect = self.calendar.rectangle() # 獲取日歷元素的矩形範圍
        today = datetime.now() # 獲取今日日期
        
        # 計算網格尺寸
        cell_width = (rect.right - rect.left) / self.GRID_COLS # 計算每個單元格的寬度
        cell_height = (rect.bottom - rect.top - self.HEADER_HEIGHT - self.WEEKDAY_HEIGHT) / self.GRID_ROWS
        
        # 計算日期位置
        first_day = today.replace(day=1) # 計算當月第一天
        first_day_col = (first_day.weekday() + 1) % 7 # 計算當月第一天是星期幾
        days_from_start = today.day - days_ago - 1 # 計算從當月第一天到今天的天數
        total_position = first_day_col + days_from_start # 計算總位置
        
        # 計算目標座標
        target_row = total_position // 7 # 計算目標行
        target_col = total_position % 7 # 計算目標列
        
        x = rect.left + (target_col + 0.5) * cell_width # 計算目標點的x座標
        y = rect.top + self.HEADER_HEIGHT + self.WEEKDAY_HEIGHT + (target_row + 0.5) * cell_height # 計算目標點的y座標
        
        return int(x), int(y)

    def click_date(self, days_ago):
        """點擊某日日期"""
        try:
            rect = self.calendar.rectangle() # 獲取日歷元素的矩形範圍
            x, y = self.calculate_click_position(days_ago) # 計算點擊位置
            if x is None or y is None:
                debug_print("無法計算點擊位置")
                return False
                
            debug_print(f"今天是 {datetime.now().day} 號") # 印出今日日期
            debug_print(f"計算得出的點擊位置: x={x}, y={y}") # 印出計算得出的點擊位置
            
            # 執行三次點擊
            for _ in range(3):
                pyautogui.click(x, y)
                time.sleep(self.CLICK_DELAY)
            
            debug_print("已執行三次點擊")
            return True
            
        except Exception as e:
            debug_print(f"點擊日期時發生錯誤: {str(e)}")
            return False
    
    def click_calendar_blank(self):
        """點擊日歷空白處"""
        try:
            rect = self.calendar.rectangle() # 獲取日歷元素的矩形範圍
            
            # 計算日歷空白處的位置 (選擇右下角)
            x = rect.left + 140 # 右邊界向右140像素
            y = rect.top + 10 # 上邊界向下10像素
            
            # 執行點擊 2 次
            pyautogui.click(x, y)
            time.sleep(self.CLICK_DELAY)
            pyautogui.click(x, y)
            time.sleep(self.CLICK_DELAY)
            
            debug_print("已點擊日歷空白處")
            return True
            
        except Exception as e:
            debug_print(f"點擊日歷空白處時發生錯誤: {str(e)}")
            return False

def start_calendar_checker(days_ago):
    """開始檢測日歷元素並點選今日日期"""
    debug_print("開始檢測日歷元素並點選今日日期...")
    
    checker = CalendarChecker() # 創建CalendarChecker實例
    if not checker.find_window():
        return
        
    if not ensure_foreground_window(checker.hwnd, checker.window_title): # 確保視窗在前景
        debug_print("警告: 無法確保視窗在前景")
    
    if not checker.find_calendar():
        debug_print("無法找到日歷元素")
        return
        
    debug_print("找到日歷元素!")
    if checker.click_date(days_ago): # 點擊日期
        try:
            debug_print(f"日歷可見性: {checker.calendar.is_visible()}")
        except Exception as e:
            debug_print(f"獲取日歷資訊時發生錯誤: {str(e)}")

def start_click_calendar_blank():
    """開始點擊日歷空白處"""
    debug_print("開始點擊日歷空白處...")

    checker = CalendarChecker()
    if not checker.find_window():
        return
    
    if not ensure_foreground_window(checker.hwnd, checker.window_title):
        debug_print("警告: 無法確保視窗在前景")
        
    if not checker.find_calendar():
        debug_print("無法找到日歷元素")
        return
    
    debug_print("找到日歷元素!")
    if checker.click_calendar_blank(): # 點擊日歷空白處
        try:
            debug_print(f"日歷可見性: {checker.calendar.is_visible()}")
        except Exception as e:
            debug_print(f"獲取日歷資訊時發生錯誤: {str(e)}")