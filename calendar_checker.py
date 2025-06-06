import win32gui
from pywinauto.application import Application
from datetime import datetime
import pyautogui
import time
from utils import (debug_print, find_window_handle, ensure_foreground_window, 
                  calculate_center_position, program_moving_context, check_mouse_before_move, click_at)
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
        
        debug_print(f"當前時間: {today.strftime('%Y-%m-%d %H:%M:%S')}", color='yellow')
        
        # 日期區域的高度（扣除標題區域）
        date_area_height = rect.bottom - rect.top - 52  # 扣除52px的標題區域
        
        # 計算網格尺寸
        cell_width = (rect.right - rect.left) / self.GRID_COLS  # 計算每個單元格的寬度
        cell_height = date_area_height / 6  # 日期區域平均分成6行
        
        # 計算目標日期
        target_date = today.replace(day=today.day - days_ago)
        raw_weekday = target_date.weekday()  # 0-6 (週一-週日)
        
        # 轉換到日曆的列序（週日=1, 週一=2, ..., 週六=7）
        if raw_weekday == 6:  # 週日
            calendar_col = 1
        else:
            calendar_col = raw_weekday + 2
            
        debug_print(f"原始 weekday(): {raw_weekday}", color='yellow')
        debug_print(f"日曆列序: {calendar_col}", color='yellow')
        
        # 計算這個日期在月曆上的位置
        first_day = target_date.replace(day=1)
        raw_first_weekday = first_day.weekday()  # 0-6 (週一-週日)
        
        # 轉換月初日期到日曆的列序
        if raw_first_weekday == 6:  # 週日
            first_calendar_col = 1
        else:
            first_calendar_col = raw_first_weekday + 2
            
        debug_print(f"月初原始 weekday(): {raw_first_weekday}", color='yellow')
        debug_print(f"月初日曆列序: {first_calendar_col}", color='yellow')
        
        # 計算從月初第一格開始的偏移天數
        days_offset = target_date.day - 1  # 從1號到目標日期的天數
        
        # 修正總偏移計算邏輯：日曆總是會在第一行顯示完整的一週
        # 無論月初是星期幾，都要加上該月1號之前的位置數
        # 這些位置會被前一個月的日期填充
        if raw_first_weekday == 6:  # 月初是週日
            # 即使是週日，前面仍有6個位置被前月日期填充
            total_offset = 6 + days_offset
        else:
            # raw_first_weekday + 1 表示前面有多少個前月日期
            total_offset = (raw_first_weekday + 1) + days_offset
        
        # 計算在第幾行（從0開始）
        target_row = total_offset // 7
        
        # 計算最終座標（使用日曆列序，需要減1因為從0開始計算）
        x = rect.left + ((calendar_col - 1) + 0.5) * cell_width  # x座標（列中心）
        y = rect.top + 52 + (target_row + 0.5) * cell_height  # y座標（從標題區域下方開始）
        
        debug_print(f"日期計算:", color='light_cyan')
        debug_print(f"目標日期: {target_date.strftime('%Y-%m-%d')}", color='light_cyan')
        debug_print(f"星期幾: {calendar_col} (1=週日, 2=週一, ..., 7=週六)", color='light_cyan')
        debug_print(f"月初星期幾: {first_calendar_col}", color='light_cyan')
        debug_print(f"日期偏移: {days_offset}", color='light_cyan')
        debug_print(f"總偏移: {total_offset}", color='light_cyan')
        debug_print(f"目標位置: 第{target_row + 1}行, 第{calendar_col}列", color='light_cyan')
        debug_print(f"點擊座標: x={int(x)}, y={int(y)}", color='light_cyan')
        
        return int(x), int(y)

    @check_mouse_before_move
    def click_date(self, days_ago):
        """點擊某日日期"""
        try:
            rect = self.calendar.rectangle()
            x, y = self.calculate_click_position(days_ago)
            if x is None or y is None:
                debug_print("無法計算點擊位置")
                return False
                
            debug_print(f"今天是 {datetime.now().day} 號", color='light_blue', bold=True)
            debug_print(f"計算得出的點擊位置: x={x}, y={y}", color='light_blue', bold=True)
            
            with program_moving_context():
                click_at(x, y, clicks=3, interval=Config.DOUBLE_CLICK_INTERVAL, 
                        hwnd=self.hwnd, window_title=self.window_title)
                
                debug_print("已執行三次點擊", color='light_yellow')
                return True
                
        except Exception as e:
            debug_print(f"點擊日期時發生錯誤: {str(e)}")
            return False

    @check_mouse_before_move
    def click_calendar_blank(self):
        """點擊日歷空白處"""
        try:
            rect = self.calendar.rectangle()
            x = rect.left + 140
            y = rect.top + 10
            
            with program_moving_context():
                click_at(x, y, clicks=2, interval=Config.DOUBLE_CLICK_INTERVAL, hwnd=self.hwnd, window_title=self.window_title)
                
                debug_print("已點擊日歷空白處", color='light_yellow')
                return True
                
        except Exception as e:
            debug_print(f"點擊日歷空白處時發生錯誤: {str(e)}")
            return False

@ensure_foreground_window
def start_calendar_checker(days_ago=0, hwnd=None, window_title=None):
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

@ensure_foreground_window
def start_click_calendar_blank(hwnd=None, window_title=None):
    """開始點擊日歷空白處"""
    try:
        debug_print("開始點擊日歷空白處...", color='light_yellow')

        checker = CalendarChecker()
        if not checker.find_window():
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