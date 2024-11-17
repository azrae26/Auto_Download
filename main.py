import keyboard
import pywinauto
import pyautogui
import time
from pywinauto.application import Application as PywinautoApp
import sys
import psutil
import win32gui
import win32con
import threading
from queue import Queue
from datetime import datetime, timedelta
from calendar_checker import start_calendar_checker, start_click_calendar_blank
from get_list_area import start_list_area_checker, set_stop, list_all_controls, monitor_clicks
from utils import (debug_print, find_window_handle, ensure_foreground_window, 
                  get_list_items_by_id, calculate_center_position, refresh_checking, 
                  start_refresh_check, stop_refresh_check, click_at, move_to_safe_position, 
                  check_mouse_movement, scroll_to_file, is_file_visible, switch_to_list)
from scheduler import Scheduler
from font_size_setter import set_font_size
from chrome_monitor import start_chrome_monitor
from folder_monitor import start_folder_monitor, FolderMonitor
from config import Config, COLORS  # 添加這行

class WindowHandler:
    """處理窗口相關操作"""
    @staticmethod
    def is_process_running(process_name):
        """檢查程序是否正在運行"""
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == process_name.lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return False

class ListNavigator:
    """處理列表導航相關操作"""
    def __init__(self):
        self.searching_up = False
        self.down_retry_count = 0
        self.up_retry_count = 0
        self.after_tab_switch = False

    def reset_search_state(self):
        self.searching_up = False
        self.down_retry_count = 0
        self.up_retry_count = 0

    def navigate_to_file(self, file, main_list_area, hwnd, file_name):
        """導航到指定檔案的位置"""
        if not self.after_tab_switch:
            self.reset_search_state()
        
        while not FileProcessor.is_file_visible(file, main_list_area):
            if self.after_tab_switch:
                if self.up_retry_count < Config.RETRY_LIMIT:
                    debug_print(f"檔案 '{file_name}' 不在可視範圍內，向上翻頁 (第 {self.up_retry_count + 1} 次)")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(Config.SLEEP_INTERVAL * 2)
                    pyautogui.press('pageup')
                    time.sleep(Config.SLEEP_INTERVAL * 5)
                    self.up_retry_count += 1
                else:
                    debug_print(f"無法在當前列表找到案 '{file_name}'，嘗試切換到下一個列表")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(Config.SLEEP_INTERVAL * 2)
                    switch_to_list(hwnd)
                    return False
            else:
                if not self.searching_up and self.down_retry_count < Config.RETRY_LIMIT:
                    debug_print(f"檔案 '{file_name}' 不在可視範圍內，向下翻頁 (第 {self.down_retry_count + 1} 次)")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(Config.SLEEP_INTERVAL * 2)
                    pyautogui.press('pagedown')
                    time.sleep(Config.SLEEP_INTERVAL * 5)
                    self.down_retry_count += 1
                elif not self.searching_up and self.down_retry_count >= Config.RETRY_LIMIT:
                    debug_print("向下找不到，開始向上翻頁尋找")
                    self.searching_up = True
                elif self.searching_up and self.up_retry_count < Config.RETRY_LIMIT:
                    debug_print(f"檔案 '{file_name}' 不在可視範圍內，向上翻頁 (第 {self.up_retry_count + 1} 次)")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(Config.SLEEP_INTERVAL * 2)
                    pyautogui.press('pageup')
                    time.sleep(Config.SLEEP_INTERVAL * 5)
                    self.up_retry_count += 1
                else:
                    debug_print(f"無法在當前列表找到檔案 '{file_name}'，嘗試切換到下一個列表")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(Config.SLEEP_INTERVAL * 2)
                    switch_to_list(hwnd)
                    return False
        return True

class FileProcessor:
    """處理文件相關操作"""
    def __init__(self):
        self.current_file_count = 0
        self.last_known_position = 0
        self.is_date_switching = False
        self.should_stop = False
        self.navigator = ListNavigator()  # 添加 ListNavigator 實例
        pyautogui.FAILSAFE = False

    @staticmethod
    def is_last_file_in_current_list(file, current_list_items):
        """判斷是否為當前列表的最後一個檔案"""
        try:
            # 獲取當前檔案在列表中的索引
            current_index = current_list_items.index(file)
            # 如果是列表中的最後一個項目
            return current_index == len(current_list_items) - 1
        except Exception as e:
            debug_print(f"檢查檔案位置時發生錯誤: {str(e)}", color='light_red')
            return False

    def close_windows(self, count):
        """關閉指定數量的chrome分頁"""
        if count <= 0:
            return
        
        time.sleep(Config.SLEEP_INTERVAL * 3)  # 等待視窗完全打開
        
        try:
            # 按下 CTRL
            keyboard.press('ctrl')
            time.sleep(Config.SLEEP_INTERVAL)  # 等待 0.1 秒 確保 CTRL 被按下
            
            # 按指定次數的 W
            for _ in range(count):
                keyboard.press('w')
                time.sleep(Config.SLEEP_INTERVAL)  # 間隔 0.1秒
                keyboard.release('w')
                time.sleep(Config.SLEEP_INTERVAL)  # 間隔 0.1秒
            
            # 釋放 CTRL
            keyboard.release('ctrl')
            time.sleep(Config.SLEEP_INTERVAL)  # 等待 0.1秒 所有視窗關閉
            
        except Exception as e:
            debug_print(f"關閉視窗時發生錯誤: {str(e)}", color='light_red')
            # 確保 CTRL 鍵被釋放
            keyboard.release('ctrl')

    def process_files(self, app, hwnd, window_title, should_stop_callback):
        """處理檔案"""
        try:
            if should_stop_callback():
                debug_print("[DEBUG] 下載開始前檢測到停止信號", color='yellow')
                return

            if not ensure_foreground_window(hwnd, window_title):
                debug_print("錯誤: 無法確保視窗可見", color='light_red')
                return

            # 連接到視窗
            main_window = app.window(handle=hwnd)
            
            # 初始化全域的已下載檔案集合
            downloaded_files = set()
            is_first_click = True
            click_count = 0
            current_list_index = 0  # 添加：追蹤當前列表索引
            
            while True:
                # 重置列表位置到第一個列表（如果不是第一次循環）
                if current_list_index > 0:
                    debug_print("重置列表位置到第一個列表", color='light_cyan')
                    for _ in range(current_list_index):
                        switch_to_list(hwnd)
                    current_list_index = 0
                
                # 獲取當前所有檔案（不含已下載的）
                all_files = []
                valid_areas = []  # 儲存有效的列表區域
                
                # 使用正確的列表類型標識符
                list_types = [
                    ('morning', '晨會報告'),
                    ('research', '研究報告'),
                    ('industry', '產業報告')
                ]
                
                # 獲取所有未下載的檔案
                for list_type, list_name in list_types:
                    debug_print(f"獲取 [{list_name}] 檔案...", color='light_cyan')
                    files = get_list_items_by_id(main_window, list_type)
                    if files:
                        # 獲取列表區域
                        list_area = main_window.child_window(auto_id=f"listBox{list_type.capitalize()}Reports")
                        valid_areas.append(list_area)
                        
                        # 只收集未下載的檔案
                        valid_files = [
                            (file, len(valid_areas) - 1) 
                            for file in files 
                            if not file.window_text().endswith("_公司") and 
                               file.window_text() not in downloaded_files
                        ]
                        
                        if valid_files:
                            all_files.extend(valid_files)
                            debug_print(f"[{list_name}] 找到 {len(valid_files)} 個未下載檔案", color='white')
                        else:
                            debug_print(f"[{list_name}] 沒有新的檔案需要下載", color='yellow')

                if not all_files:
                    debug_print("所有檔案已下載完成", color='light_green')
                    break

                debug_print(f"本輪需要下載 {len(all_files)} 個檔案", color='light_green')

                # 下載檔案
                for file, area_index in all_files:
                    if should_stop_callback(): 
                        return

                    try:
                        file_name = file.window_text()
                        list_area = valid_areas[area_index]
                        
                        # 如果需要切換到不同的列表
                        if area_index != current_list_index:
                            debug_print(f"切換列表: 從 {current_list_index} 到 {area_index}", color='light_cyan')
                            for _ in range(area_index - current_list_index):
                                switch_to_list(hwnd)
                            current_list_index = area_index
                            time.sleep(Config.SLEEP_INTERVAL * 2)  # 等待列表切換完成
                        
                        # 檢查檔案可見性
                        if not is_file_visible(file, list_area):
                            debug_print(f"檔案 '{file_name}' 不在可視範圍內，嘗試調整位置", color='light_magenta')
                            if not scroll_to_file(file, list_area, hwnd):
                                debug_print(f"無法使檔案 '{file_name}' 進入可視範圍，跳過")
                                continue
                        
                        # 執行點擊
                        rect = file.rectangle()
                        center_x, center_y = calculate_center_position(rect)
                        if center_x is None or center_y is None:
                            debug_print("無法計算檔案位置，跳過此檔案")
                            continue
                        
                        click_at(
                            center_x, 
                            center_y, 
                            clicks=2, 
                            interval=Config.SLEEP_INTERVAL * 0.5, 
                            is_first_click=is_first_click, 
                            hwnd=hwnd, 
                            window_title=window_title,
                            expected_text=file_name
                        )
                        debug_print(f"下載完成: {file_name}", color='white')
                        
                        if not is_first_click:
                            click_count += 1
                            if click_count == Config.CLICK_BATCH_SIZE:
                                self.close_windows(Config.CLICK_BATCH_SIZE)
                                click_count = 0
                        else:
                            is_first_click = False
                            time.sleep(Config.SLEEP_INTERVAL * 5)

                        downloaded_files.add(file_name)
                        
                        if self.is_last_file_in_list(file, area_index, all_files):
                            if area_index == len(valid_areas) - 1:
                                debug_print(f"已經是最後一個列表，跳過切換", color='yellow')
                            else:
                                debug_print(f"切換到下一個列表", color='light_cyan')
                                switch_to_list(hwnd)
                                current_list_index += 1  # 更新當前列表索引

                    except Exception as e:
                        debug_print(f"處理檔案時發生錯誤: {str(e)}", color='light_red')
                        continue

                # 關閉剩餘的Chrome視窗
                if click_count > 0:
                    self.close_windows(click_count)

                # 檢查是否有漏掉的檔案
                debug_print("檢查是否有漏掉的檔案...", color='light_cyan')
                
                # 獲取所有檔案（不含_公司）
                current_files = set()
                for list_type, _ in list_types:
                    files = get_list_items_by_id(main_window, list_type)
                    current_files.update(
                        file.window_text() for file in files 
                        if not file.window_text().endswith("_公司")
                    )

                # 找出漏掉的檔案
                missed_files = current_files - downloaded_files

                if not missed_files:  # 如果沒有漏掉的檔案，跳出循環
                    debug_print("沒有漏掉的檔案，下載完成", color='light_green')
                    break
                else:
                    debug_print(f"發現 {len(missed_files)} 個漏掉的檔案，繼續下載...", color='light_cyan')
                    # 下一輪循環將只處理漏掉的檔案

            debug_print("所有檔案下載完成", color='light_green')

        except Exception as e:
            debug_print(f"處理檔案時發生錯誤: {str(e)}", color='light_red')

    def is_last_file_in_list(self, current_file, list_index, all_files):
        """判斷是否為當前列表的最後一個檔案"""
        current_index = all_files.index((current_file, list_index))
        if current_index + 1 >= len(all_files):
            return True
        next_file, next_list_index = all_files[current_index + 1]
        return next_list_index != list_index

    def get_all_files(self, main_window):
        """獲取所有檔案"""
        all_files = []
        for i in range(3):  # 三個列表
            files = get_list_items_by_id(main_window, i)
            all_files.extend([(file, i) for file in files if not file.window_text().endswith("_公司")])
        return all_files

class MainApp:
    """主應用程序類"""
    def __init__(self):
        self.file_processor = FileProcessor()
        self.selected_window = None
        self.scheduler = None
        self.should_stop = False
        self.stop_event = threading.Event()
        self.esc_thread = None  # 新增：保存 ESC 監聽線程的引用
        self.chrome_monitor = None
        self.collected_lists = {
            '今日': [], '昨日': [], '1週前': [], '2週前': [], '3週前': [], 
            '4週前': [], '5週前': [], '6週前': [], '7週前': [], '8週前': []
        }

    def check_esc_key(self):
        """監聽 ESC 按鍵"""
        while not self.stop_event.is_set():  # 使用 Event 來控制線程
            if keyboard.is_pressed('esc'):
                self.should_stop = True
                self.stop_event.set()  # 設置停止事件
                debug_print("\n[DEBUG] ESC 按鍵被按下，停止執行")
                # 同步停止其他模組
                set_stop()  # 停止 get_list_area
                break
            time.sleep(0.1)

    def start_esc_listener(self):
        """啟動 ESC 監聽"""
        if self.esc_thread is None or not self.esc_thread.is_alive():
            self.should_stop = False
            self.stop_event.clear()
            self.esc_thread = threading.Thread(target=self.check_esc_key, daemon=True)
            self.esc_thread.start()

    def stop_esc_listener(self):
        """停止 ESC 監聽"""
        self.stop_event.set()
        if self.esc_thread and self.esc_thread.is_alive():
            self.esc_thread.join(timeout=1.0) # 等待1秒，確保線程結束

    def select_window(self, index):
        """選擇視窗並處理檔案"""
        self.start_esc_listener()  # 使用新的啟動方法
        
        # 獲取視窗句柄
        target_windows = find_window_handle(Config.TARGET_WINDOW)
        if 1 <= index <= len(target_windows):
            # 選擇視窗
            self.selected_window = target_windows[index - 1]
            debug_print(f"\n已選擇視窗: {self.selected_window[1]}")
            hwnd, window_title = self.selected_window
            # 連接到視窗
            app = PywinautoApp(backend="uia").connect(handle=hwnd)            
            # 處理檔案
            self.file_processor.process_files(app, hwnd, window_title, lambda: self.should_stop)
        else:
            debug_print("無效的選擇")
        
        self.stop_esc_listener()  # 操作完成後停止監聽

    def click_daily_report_tab(self, hwnd=None, window_title=None):
        """點擊每日報告標籤"""
        try:
            # 如果沒有提供視窗句柄，則獲取
            if hwnd is None or window_title is None:
                hwnd, window_title = self.selected_window or find_window_handle(Config.TARGET_WINDOW)[0]
            
            # 確保視窗在前景
            if not ensure_foreground_window(hwnd, window_title):
                debug_print("警告: 無法確保視窗前景", color='light_red')
                return False
                
            # 連接到視窗
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
            main_window = app.window(handle=hwnd)
            
            # 找到每日報告標籤
            daily_report_tab = main_window.child_window(title="每日報告", control_type="TabItem")
            
            # 獲取位置
            rect = daily_report_tab.rectangle() # 獲取矩形
            center_x, center_y = calculate_center_position(rect) # 計算中心位置
            
            if center_x and center_y:
                # 雙擊
                click_at(center_x, center_y, clicks=2, interval=Config.SLEEP_INTERVAL, 
                                          hwnd=hwnd, window_title=window_title)
                debug_print("已點擊每日報告標籤", color='yellow')
                return True
            else:
                debug_print("無法獲取每日報告標籤位置")
                return False
                
        except Exception as e:
            debug_print(f"點擊每日報告標籤時發生錯誤: {str(e)}", color='light_red')
            return False

    def execute_sequence(self):
        """執行連續任務"""
        self.start_esc_listener()
        debug_print("開始執行連續任務...", color='light_cyan')
        
        # 先獲取視窗句柄
        target_windows = find_window_handle(Config.TARGET_WINDOW)
        if not target_windows:
            debug_print("錯誤: 找不到目標視窗", color='light_red')
            return
        
        hwnd, window_title = target_windows[0]
        folder_monitor = FolderMonitor()
        
        # 使用 self.collected_lists 而不是創建新的字典
        self.collected_lists = {
            '今日': [], '昨日': [], '1週前': [], '2週前': [], '3週前': [], 
            '4週前': [], '5週前': [], '6週前': [], '7週前': [], '8週前': []
        }
        
        def press_left_or_up(left_times, up_times):
            """連續按左或上鍵指定次數"""
            if left_times > 0:
                debug_print(f"按下左鍵 {left_times} 次", color='yellow')
            for _ in range(left_times):
                pyautogui.press('left')
                time.sleep(Config.SLEEP_INTERVAL)
            
            if up_times > 0:
                debug_print(f"按下上鍵 {up_times} 次", color='yellow')
            for _ in range(up_times):
                pyautogui.press('up')
                time.sleep(Config.SLEEP_INTERVAL)
        
        def download_days_weeks(days_ago, weeks_ago, list_name=None):
            """下載 N 天前、或 N 週前的檔案，並收集列表"""
            if days_ago > 0 or weeks_ago > 0:
                press_left_or_up(days_ago, weeks_ago)
            
            # 在下載前收集當前列表
            if list_name and hwnd:
                self.collect_current_list(list_name, hwnd)
            
            self.select_window(1)
        
        # 基本步驟
        steps = [
            ("點擊每日報告標籤", lambda: self.click_daily_report_tab(hwnd, window_title)),
            ("設定字型大小", lambda: set_font_size()),
            ("點擊今日", lambda: start_calendar_checker(0)),
            ("下載今日檔案，並收集列表", lambda: download_days_weeks(0, 0, '今日')),
            ("點擊今日", lambda: start_calendar_checker(0)),
            ("下載昨日檔案，並收集列表", lambda: download_days_weeks(1, 0, '昨日')),
            ("點擊今日", lambda: start_calendar_checker(0)),
            ("下載 1 週前檔案，並收集列表", lambda: download_days_weeks(0, 1, '1週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank()),
            ("下載 2 週前檔案，並收集列表", lambda: download_days_weeks(0, 1, '2週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank()),
            ("下載 4 週前檔案，並收集列表", lambda: download_days_weeks(0, 2, '4週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank()),
            ("下載 8 週前檔案，並收集列表", lambda: download_days_weeks(0, 4, '8週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank()),
            ("鍵盤向下 X8", lambda: [pyautogui.press('down') or time.sleep(0.2) for _ in range(8)]),
            ("點擊今日", lambda: start_calendar_checker(0)),
            ("複製今日檔案到指定位置", lambda: self.copy_today_files()),
            ("分析檔案匹配", lambda: folder_monitor.store_and_analyze_lists(
                today_list=self.collected_lists['今日'],
                yesterday_list=self.collected_lists['昨日'],
                last_week_list=self.collected_lists['1週前'],
                last_2week_list=self.collected_lists['2週前'],
                last_3week_list=self.collected_lists['4週前'],
                last_4week_list=self.collected_lists['8週前']
            ))
        ]
        
        # 執行所有步驟
        for i, (step_name, step_func) in enumerate(steps, 1):
            if self.should_stop:
                debug_print("任務已停止", color='yellow')
                break
            
            debug_print(f"步驟{i}: {step_name}", color='yellow')
            step_func()
            time.sleep(1)
        
        debug_print("連續任務執行完成", color='light_green')
        self.stop_esc_listener()

    def download_current_list(self):
        """下載當前列表檔案"""
        self.should_stop = False
        self.stop_event.clear()
        
        debug_print("開始下載當前列表檔案...", color='light_cyan')
        
        self.start_esc_listener()
        move_to_safe_position()
        self.select_window(1)

    def toggle_refresh_check(self):
        """切換列表刷新檢測"""
        if not refresh_checking:
            hwnd = self.selected_window[0] if self.selected_window else None
            debug_print("開始檢測列表刷新", color='light_magenta')
            threading.Thread(
                target=start_refresh_check,
                args=(hwnd,),
                daemon=True
            ).start()
        else:
            debug_print("停止檢測列表刷新", color='yellow')
            stop_refresh_check()

    def monitor_chrome(self):
        """切換 Chrome 監控狀態"""
        self.chrome_monitor = start_chrome_monitor(self.chrome_monitor)

    def copy_today_files(self):
        """複製今日檔案的快捷鍵處理"""
        folder_monitor = FolderMonitor()
        folder_monitor.copy_today_files()

    def list_all_reports(self):
        """列出所有報告清單"""
        try:
            # 獲取視窗句柄
            windows = find_window_handle(Config.TARGET_WINDOW)
            if not windows:
                debug_print("找不到目標視窗", color='light_red')
                return
            
            hwnd = windows[0][0]
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
            main_window = app.window(handle=hwnd)
            
            # 列出三個列表的內容
            list_types = {
                'morning': '晨會報告',
                'research': '研究報告',
                'industry': '產業報告'
            }
            
            for list_type, list_name in list_types.items():
                debug_print(f"\n=== {list_name} ===", color='light_cyan')
                items = get_list_items_by_id(main_window, list_type)
                for i, item in enumerate(items, 1):
                    debug_print(f"{i}. {item.window_text()}", color='light_green')
                
        except Exception as e:
            debug_print(f"列出報告清單時發生錯誤: {str(e)}", color='light_red')

    def collect_current_list(self, list_name, hwnd):
        """
        從指定窗口收集列表項目
        
        Args:
            list_name (str): 列表名稱
            hwnd: 窗口句柄
        
        Returns:
            bool: 收集是否成功
        """
        if list_name and hwnd:
            try:
                app = PywinautoApp(backend="uia").connect(handle=hwnd)
                main_window = app.window(handle=hwnd)
                self.collected_lists[list_name] = []
                for list_type in ['morning', 'research', 'industry']:
                    files = get_list_items_by_id(main_window, list_type)
                    if files:
                        self.collected_lists[list_name].extend([f.window_text() for f in files if f.window_text()])
                debug_print(f"已收集 {list_name} 列表，共 {len(self.collected_lists[list_name])} 個檔案", color='light_green')
                return True
            except Exception as e:
                debug_print(f"收集列表時發生錯誤: {str(e)}", color='light_red')
                return False
        return False

    def collect_and_analyze_lists(self):
        """收集各時間點列表並分析"""
        debug_print("\n開始收集列表...", color='light_cyan')
        
        # 先獲取視窗句柄
        target_windows = find_window_handle(Config.TARGET_WINDOW)
        if not target_windows:
            debug_print("錯誤: 找不到目標視窗", color='light_red')
            return
        
        hwnd, window_title = target_windows[0]
        folder_monitor = FolderMonitor()
        
        # 執行步驟
        steps = [
            ("等待 1 秒", lambda: time.sleep(1)),
            ("點擊每日報告標籤", lambda: self.click_daily_report_tab(hwnd, window_title)),
            ("點擊今日", lambda: start_calendar_checker(0) or time.sleep(2)),
            ("收集今日列表", lambda: self.collect_current_list('今日', hwnd)),
            ("向左 1 天", lambda: pyautogui.press('left') or time.sleep(2)),
            ("收集昨日列表", lambda: self.collect_current_list('昨日', hwnd)),
            ("點擊今日", lambda: start_calendar_checker(0) or time.sleep(2)),
            ("向上 1 週", lambda: pyautogui.press('up') or time.sleep(2)),
            ("收集 1 週前列表", lambda: self.collect_current_list('1週前', hwnd)),
            ("向上 1 週", lambda: [pyautogui.press('up') or time.sleep(2) for _ in range(1)]),
            ("收集 2 週前列表", lambda: self.collect_current_list('2週前', hwnd)),
            ("向上 2 週", lambda: [pyautogui.press('up') or time.sleep(2) for _ in range(2)]),
            ("收集 4 週前列表", lambda: self.collect_current_list('4週前', hwnd)),
            ("向上 4 週", lambda: [pyautogui.press('up') or time.sleep(2) for _ in range(4)]),
            ("收集 8 週前列表", lambda: self.collect_current_list('8週前', hwnd)),
            ("鍵盤向下 X8", lambda: [pyautogui.press('down') or time.sleep(2) for _ in range(8)]),
            ("點擊今日", lambda: start_calendar_checker(0)),
            ('分析檔案匹配', lambda: folder_monitor.store_and_analyze_lists(
                today_list=self.collected_lists['今日'],
                yesterday_list=self.collected_lists['昨日'],
                last_week_list=self.collected_lists['1週前'],
                last_2week_list=self.collected_lists['2週前'],
                last_3week_list=self.collected_lists['4週前'],
                last_4week_list=self.collected_lists['8週前']
            ))
        ]
        
        # 執行步驟
        for step_name, step_func in steps:
            debug_print(f"執行: {step_name}", color='yellow')
            step_func()
        
    def run(self):
        try:
            debug_print("=== 研究報告自動下載程式 ===", color='light_blue', bold=True)
            
            # 註冊所有熱鍵
            keyboard.add_hotkey('ctrl+shift+e', self.execute_sequence)
            keyboard.add_hotkey('ctrl+shift+f', self.download_current_list)
            keyboard.add_hotkey('ctrl+shift+g', start_list_area_checker)
            keyboard.add_hotkey('ctrl+shift+b', set_font_size)
            keyboard.add_hotkey('ctrl+shift+t', self.toggle_refresh_check)
            keyboard.add_hotkey('ctrl+shift+r', list_all_controls)
            keyboard.add_hotkey('ctrl+shift+m', monitor_clicks)
            keyboard.add_hotkey('ctrl+shift+k', self.monitor_chrome)
            keyboard.add_hotkey('ctrl+shift+f12', self.copy_today_files)
            keyboard.add_hotkey('ctrl+shift+f11', self.list_all_reports)
            keyboard.add_hotkey('ctrl+shift+f10', self.collect_and_analyze_lists)  # 新增這行
            
            # 顯示熱鍵說明
            debug_print("=== 快捷鍵說明 ===", color='light_cyan')
            debug_print("按下 CTRL + SHIFT + E    開始連續下載任務", color='light_green')
            debug_print("按下 CTRL + SHIFT + F    下載當前列表檔案", color='light_green')
            debug_print("按下 CTRL + SHIFT + G    檢測檔案列表區域", color='light_green')
            debug_print("按下 CTRL + SHIFT + B    設定字型大小", color='light_green')
            debug_print("按下 CTRL + SHIFT + T    切換列表刷新檢測", color='light_green')
            debug_print("按下 CTRL + SHIFT + R    列出所有控件", color='light_green')
            debug_print("按下 CTRL + SHIFT + M    開始監控滑鼠點擊", color='light_green')
            debug_print("按下 CTRL + SHIFT + K    監控 Chrome 視窗", color='light_green')
            debug_print("按下 CTRL + SHIFT + F11  列出所有報告清單", color='light_green')
            debug_print("按下 CTRL + SHIFT + F12  複製今日所有新檔案", color='light_green')
            debug_print("按下 CTRL + SHIFT + F10  收集並分析所有列表", color='light_green')
            debug_print("按下 ESC                 停止下載", color='yellow')
            debug_print("按下 CTRL+SHIFT+Q        關閉程式", color='light_red', bold=True)
            debug_print("==================", color='light_cyan')

            self.scheduler = Scheduler(self.execute_sequence)
            scheduler_thread = self.scheduler.init_scheduler() # 初始化排程器
            schedule_times = Config.get_schedule_times() # 獲取排程時間
            
            # 使用阻塞方式等待 Ctrl+Shift+Q
            try:
                keyboard.wait('ctrl+shift+q')
                debug_print("收到關閉程式的命令", color='light_magenta')
            except Exception as e:
                debug_print(f"等待關閉命令時發生錯誤: {str(e)}", color='light_red')
            
        except Exception as e:
            debug_print(f"程式執行時發生錯誤: {str(e)}", color='light_red')
        finally:
            debug_print("正在清理並關閉程式...", color='light_cyan')
            keyboard.unhook_all()

def main():
    app = MainApp()  # 創建主應用程式實例
    
    try:
        app.run()  # 運行主應用程式
    except:  # 捕獲所有異常但不處理
        pass

if __name__ == "__main__":
    main()
