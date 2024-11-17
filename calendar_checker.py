import win32gui
from pywinauto.application import Application
from datetime import datetime
import pyautogui
import time
from utils import (debug_print, find_window_handle, ensure_foreground_window, 
                  calculate_center_position, program_moving_context)
from config import Config, COLORS
from control_info import get_calendar_info

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
        windows = find_window_handle(Config.TARGET_WINDOW)
        
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
            
            # 使用 control_info.py 中定義的日曆控件資訊
            calendar_info = get_calendar_info()
            
            # 先找到所有符合條件的控件
            candidates = main_window.descendants(
                class_name=calendar_info['class_name'],
                control_type=calendar_info['element_info']['control_type']
            )
            
            debug_print(f"找到 {len(candidates)} 個候選控件", color='light_cyan')
            
            # 遍歷所有候選控件，找到正確的日曆控件
            for candidate in candidates:
                try:
                    rect = candidate.rectangle()
                    width = rect.right - rect.left
                    height = rect.bottom - rect.top
                    
                    # 檢查尺寸是否符合（允許5像素的誤差）
                    if (abs(width - 186) <= 5 and abs(height - 162) <= 5):
                        debug_print(f"檢查控件:", color='light_cyan')
                        debug_print(f"寬度: {width}, 高度: {height}", color='light_cyan')
                        debug_print(f"類型: {candidate.element_info.control_type}", color='light_cyan')
                        debug_print(f"類別: {type(candidate).__name__}", color='light_cyan')
                        debug_print(f"類別名稱: {candidate.class_name()}", color='light_cyan')
                        
                        # 找到符合的控件
                        self.calendar = candidate
                        debug_print("找到符合的日曆控件!", color='light_green')
                        return True
                        
                except Exception as e:
                    continue
            
            debug_print("無法找到符合尺寸的日曆控件", color='light_red')
            return False
            
        except Exception as e:
            debug_print(f"尋找日曆元素時發生錯誤: {str(e)}", color='light_red')
            return False

    def calculate_click_position(self, days_ago):
        """計算點擊位置"""
        rect = self.calendar.rectangle()  # 獲取日歷元素的矩形範圍
        today = datetime.now()  # 獲取今日日期
        
        # 日期區域的高度（扣除標題區域）
        date_area_height = rect.bottom - rect.top - 52  # 扣除52px的標題區域
        
        # 計算網格尺寸（確保只有6行）
        cell_width = (rect.right - rect.left) / self.GRID_COLS  # 計算每個單元格的寬度
        cell_height = date_area_height / 6  # 日期區域平均分成6行
        
        # 計算日期位置
        first_day = today.replace(day=1)  # 計算當月第一天
        first_day_col = (first_day.weekday() + 1) % 7  # 計算當月第一天是星期幾
        days_from_start = today.day - days_ago - 1  # 計算從當月第一天到今天的天數
        total_position = first_day_col + days_from_start  # 計算總位置
        
        # 計算目標座標
        target_row = total_position // 7  # 計算目標行（0-5）
        target_col = total_position % 7  # 計算目標列（0-6）
        
        # 計算最終座標
        x = rect.left + (target_col + 0.5) * cell_width  # x座標（列中心）
        y = rect.top + 52 + (target_row + 0.5) * cell_height  # y座標（從標題區域下方開始）
        
        debug_print(f"日曆區域: 左={rect.left}, 上={rect.top}, 右={rect.right}, 下={rect.bottom}", color='light_cyan')
        debug_print(f"日歷尺寸: 寬={rect.right - rect.left}px, 高={rect.bottom - rect.top}px", color='light_cyan')
        debug_print(f"日期區高度: {date_area_height}px, 寬度: {rect.right - rect.left}px", color='light_cyan')
        debug_print(f"單元格尺寸: 寬={cell_width:.1f}px, 高={cell_height:.1f}px", color='light_cyan')
        debug_print(f"目標位置: 第{target_row + 1}行, 第{target_col + 1}列", color='light_cyan')
        debug_print(f"點擊座標: x={int(x)}, y={int(y)}", color='light_cyan')
        
        return int(x), int(y)

    def click_date(self, days_ago):
        """點擊某日日期"""
        try:
            rect = self.calendar.rectangle() # 獲取日歷元素的矩形範圍
            x, y = self.calculate_click_position(days_ago) # 計算點擊位置
            if x is None or y is None:
                debug_print("無法計算點擊位置")
                return False
                
            debug_print(f"今天是 {datetime.now().day} 號", color='light_blue', bold=True) # 印出今日日期
            debug_print(f"計算得出的點擊位置: x={x}, y={y}", color='light_blue', bold=True) # 印出計算得出的點擊位置
            
            # 只在實際點擊時使用上下文管理器
            with program_moving_context():
                # 執行三次點擊
                for _ in range(3):
                    pyautogui.click(x, y)
                    time.sleep(self.CLICK_DELAY)
                
                debug_print("已執行三次點擊", color='yellow')
                return True
                
        except Exception as e:
            debug_print(f"點擊日期時發生錯誤: {str(e)}")
            return False
    
    def click_calendar_blank(self):
        """點擊日歷空白處"""
        global is_program_moving  # 添加這行
        
        try:
            rect = self.calendar.rectangle() # 獲取日歷元素的矩形範圍
            
            # 計算日歷空白處的位置 (選擇右下角)
            x = rect.left + 140 # 右邊界向右140像素
            y = rect.top + 10 # 上邊界向下10像素
            
            # 標記為程式移動
            is_program_moving = True
            
            try:
                # 執行點擊 2 次
                pyautogui.click(x, y)
                time.sleep(self.CLICK_DELAY)
                pyautogui.click(x, y)
                time.sleep(self.CLICK_DELAY)
                
                debug_print("已點擊日歷空白處", color='yellow')
                return True
                
            finally:
                is_program_moving = False  # 確保標記被重置
                
        except Exception as e:
            debug_print(f"點擊日歷空白處時發生錯誤: {str(e)}")
            is_program_moving = False  # 確保發生錯誤時重設標記
            return False

def start_calendar_checker(days_ago=0):
    """開始檢測日歷元素並點選日期"""
    try:
        debug_print("開始檢測日歷元素並點選今日日期...", color='light_cyan')
        
        # 獲取目標視窗
        target_windows = find_window_handle(Config.TARGET_WINDOW)
        if not target_windows:
            debug_print("找不到目標視窗", color='light_red')
            return False
            
        hwnd = target_windows[0][0]
        window_title = target_windows[0][1]
        
        # 初始化 CalendarChecker 實例
        checker = CalendarChecker()
        if not checker.find_window():
            return False
        
        if not ensure_foreground_window(hwnd, window_title):
            debug_print("警告: 無法確保視窗在前景")
            return False
        
        # 找到日歷元素
        if not checker.find_calendar():
            debug_print("無法找到日歷元素")
            return False
        
        debug_print("找到日歷元素!", color='light_blue', bold=True)

        # 點擊日期 - 只有這部分需要使用上下文管理器
        with program_moving_context():
            if checker.click_date(days_ago):
                try:
                    debug_print(f"日歷可見性: {checker.calendar.is_visible()}", color='light_blue', bold=True)
                    return True
                except Exception as e:
                    debug_print(f"獲取日歷資訊時發生錯誤: {str(e)}")
                    return False
            return False
            
    except Exception as e:
        debug_print(f"檢測日歷元素時發生錯誤: {str(e)}", color='light_red')
        return False

def start_click_calendar_blank():
    """開始點擊日歷空白處"""
    try:
        debug_print("開始點擊日歷空白處...", color='yellow')

        checker = CalendarChecker()
        if not checker.find_window():
            return False
        
        if not ensure_foreground_window(checker.hwnd, checker.window_title):
            debug_print("警告: 無法確保視窗在前景")
            return False
            
        if not checker.find_calendar():
            debug_print("無法找到日歷元素")
            return False
        
        debug_print("找到日歷元素!", color='light_blue', bold=True)
        if checker.click_calendar_blank(): # 點擊日歷空白處
            try:
                debug_print(f"日歷可見性: {checker.calendar.is_visible()}", color='light_blue', bold=True)
                return True
            except Exception as e:
                debug_print(f"獲取日歷資訊時發生錯誤: {str(e)}")
                return False
        return False
        
    except Exception as e:
        debug_print(f"點擊日歷空白處時發生錯誤: {str(e)}", color='light_red')
        return False