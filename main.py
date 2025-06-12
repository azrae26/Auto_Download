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
                  check_mouse_movement, scroll_to_file, is_file_visible, switch_to_list, 
                  reset_mouse_position, check_mouse_before_move)
from scheduler import Scheduler
from font_size_setter import set_font_size
from chrome_monitor import start_chrome_monitor
from folder_monitor import start_folder_monitor, FolderMonitor
from config import Config, COLORS  # 添加這行
from test_terminal import test_terminal_support

class FileProcessor:
    """
    處理文件相關操作
    功能：自動下載檔案、關閉Chrome分頁
    職責：檔案下載流程控制、Chrome視窗管理
    依賴：win32gui（視窗檢測）、keyboard（鍵盤操作）、pyautogui（滑鼠操作）
    
    新增功能：智能關閉Chrome分頁
    - 每次關閉前檢測Chrome視窗標題
    - 關閉後驗證視窗是否真的關閉
    - 失敗時自動重試（最多3次）
    - 返回實際關閉的分頁數量
    """
    def __init__(self):
        self.current_file_count = 0
        self.last_known_position = 0
        self.is_date_switching = False
        self.should_stop = False
        pyautogui.FAILSAFE = False

    def get_chrome_window_titles(self):
        """獲取所有Chrome視窗標題"""
        def callback(hwnd, titles):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if 'chrome' in title.lower() or '.pdf' in title.lower():
                    titles.append(title)
        
        titles = []
        win32gui.EnumWindows(callback, titles)
        return titles

    def close_windows(self, count, initial_window_titles=None):
        """關閉指定數量的chrome分頁"""
        if count <= 0:
            return 0
                
        time.sleep(Config.CLOSE_WINDOW_INTERVAL * 6)
        successfully_closed = 0
        
        try:
            if count > 0:
                # 按住 Ctrl 一次就好
                keyboard.press('ctrl')
                
                for i in range(count):
                    # 記錄關閉前的所有Chrome視窗標題
                    before_close_titles = set(self.get_chrome_window_titles())
                    if not before_close_titles:
                        break
                    
                    # 只需要按 W
                    keyboard.press('w')
                    time.sleep(Config.CLOSE_WINDOW_INTERVAL / 3)
                    keyboard.release('w')
                    
                    # 等待Chrome反應，然後檢查標題變化
                    start_time = time.time()
                    title_changed = False
                    while time.time() - start_time < 3.0:  # 等待最多3秒
                        current_titles = set(self.get_chrome_window_titles())
                        if current_titles != before_close_titles:
                            # 找出變化的標題
                            removed_titles = before_close_titles - current_titles
                            added_titles = current_titles - before_close_titles
                            
                            if removed_titles:
                                debug_print(f"消失的標題: {list(removed_titles)}", color='light_red')
                            if added_titles:
                                debug_print(f"新增的標題: {list(added_titles)}", color='light_green')
                                
                            successfully_closed += 1
                            title_changed = True
                            break
                        time.sleep(0.03)  # 每0.03秒檢查一次
                    
                    if not title_changed:
                        # 3秒後視窗標題沒變，視為關閉成功
                        successfully_closed += 1
                
                # 檢查是否回到初始狀態，如果沒有則繼續關閉（最多額外關2個）
                if initial_window_titles is not None:
                    extra_close_count = 0
                    max_extra_close = 2
                    
                    while extra_close_count < max_extra_close:
                        current_titles = set(self.get_chrome_window_titles())
                        # 檢查是否有新增的視窗（相比初始狀態）
                        new_windows = current_titles - initial_window_titles
                        
                        if not new_windows:
                            debug_print("已回到初始Chrome視窗狀態，關閉完成", color='light_green')
                            break
                        
                        debug_print(f"仍有 {len(new_windows)} 個新視窗未關閉，額外關閉第 {extra_close_count + 1} 個", color='light_yellow')
                        debug_print(f"新視窗: {list(new_windows)}", color='light_cyan')
                        
                        # 額外關閉一個分頁
                        keyboard.press('w')
                        time.sleep(Config.CLOSE_WINDOW_INTERVAL / 3)
                        keyboard.release('w')
                        time.sleep(Config.CLOSE_WINDOW_INTERVAL)
                        
                        extra_close_count += 1
                        successfully_closed += 1
                    
                    if extra_close_count >= max_extra_close:
                        remaining_new = set(self.get_chrome_window_titles()) - initial_window_titles
                        debug_print(f"已額外關閉 {max_extra_close} 個分頁，停止關閉", color='light_yellow')
                        if remaining_new:
                            debug_print(f"仍有 {len(remaining_new)} 個新視窗未關閉", color='light_red')
                
                # 最後才釋放 Ctrl
                keyboard.release('ctrl')
                
        except Exception as e:
            debug_print(f"關閉視窗時發生錯誤: {str(e)}", color='light_red')
            try:
                keyboard.release('ctrl')
            except:
                pass
        
        return successfully_closed

    def process_files(self, app, hwnd, window_title, should_stop_callback):
        """處理檔案"""
        try:
            if should_stop_callback():
                debug_print("[DEBUG] 下載開始前檢測到停止信號", color='light_yellow')
                return

            if not ensure_foreground_window(hwnd, window_title):
                debug_print("錯誤: 無法確保視窗可見", color='light_red')
                return

            # 連接到視窗
            main_window = app.window(handle=hwnd)
            
            # 應該在這裡也重置滑鼠位置記錄
            reset_mouse_position()  # 添加這行
            
            # 預加載所有列表區域和檔案信息
            list_areas = {}
            list_files = {}
            list_types = [
                ('morning', '晨會報告'),
                ('research', '研究報告'),
                ('industry', '產業報告')
            ]
            
            debug_print("開始預加載列表信息...", color='light_cyan')
            for list_type, list_name in list_types:
                try:
                    # 預加載列表區域
                    list_areas[list_type] = main_window.child_window(
                        auto_id=f"listBox{list_type.capitalize()}Reports"
                    )
                    # 預加載檔案信息
                    list_files[list_type] = get_list_items_by_id(main_window, list_type)
                    debug_print(f"已預加載 [{list_name}] 列表，共 {len(list_files[list_type])} 個檔案", color='light_green')
                except Exception as e:
                    debug_print(f"預加載 {list_name} 列表時發生錯誤: {str(e)}", color='light_red')
            
            # 初始化全域的已下載檔案集合
            downloaded_files = set()
            is_first_click = True
            click_count = 0
            current_list_index = 0
            
            # 記住初始Chrome視窗標題集合（用於關閉分頁驗證）
            initial_window_titles = set(self.get_chrome_window_titles())
            debug_print(f"記住初始Chrome視窗 {len(initial_window_titles)} 個: {list(initial_window_titles)}", color='light_cyan')
            
            while True:
                # 重置列表位置到第一個列表（如果不是第一次循環）
                if current_list_index > 0:
                    debug_print("重置列表位置到第一個列表", color='light_cyan')
                    for _ in range(current_list_index):
                        switch_to_list(hwnd)
                    current_list_index = 0
                
                # 使用預加載的數據獲取所有未下載的檔案
                all_files = []
                valid_areas = []
                
                for list_type, list_name in list_types:
                    if list_type in list_files and list_files[list_type]:
                        # 使用預加載的列表區域
                        valid_areas.append(list_areas[list_type])
                        
                        # 使用預加載的檔案信息
                        valid_files = [
                            (file, len(valid_areas) - 1) 
                            for file in list_files[list_type]
                            if not file.window_text().endswith("_公司") and 
                               file.window_text() not in downloaded_files
                        ]
                        
                        if valid_files:
                            all_files.extend(valid_files)
                            debug_print(f"[{list_name}] 找到 {len(valid_files)} 個未下載檔案", color='white')
                        else:
                            debug_print(f"[{list_name}] 沒有新的檔案需要下載", color='light_yellow')
                
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
                        
                        # 執行點擊並檢查結果
                        rect = file.rectangle()
                        center_x, center_y = calculate_center_position(rect)
                        if center_x is None or center_y is None:
                            debug_print("無法計算檔案位置，跳過此檔案")
                            continue
                        
                        if click_at(
                            center_x, 
                            center_y, 
                            clicks=2, 
                            interval=Config.DOUBLE_CLICK_INTERVAL,
                            sleep_interval=Config.DOWNLOAD_INTERVAL,
                            is_first_click=is_first_click, 
                            hwnd=hwnd, 
                            window_title=window_title,
                            expected_text=file_name
                        ):  # 只有在點擊成功時才執行後續操作
                            debug_print(f"下載完成: {file_name}", color='white')
                            
                            # 每次成功點擊都計入
                            if is_first_click:
                                is_first_click = False
                            
                            click_count += 1
                            # 如果達到批次大小，關閉視窗
                            if click_count >= Config.CLICK_BATCH_SIZE:
                                self.close_windows(Config.CLICK_BATCH_SIZE, initial_window_titles)
                                click_count = 0

                            downloaded_files.add(file_name)  # 只有在點擊成功時才加入下載清單
                        else:
                            debug_print(f"下載失敗: {file_name}", color='light_red')
                            continue  # 如果點擊失敗，跳過後續操作

                        if self.is_last_file_in_list(file, area_index, all_files):
                            if area_index == len(valid_areas) - 1:
                                debug_print(f"已經是最後一個列表，跳過切換", color='light_yellow')
                            else:
                                debug_print(f"切換到下一個列表", color='light_cyan')
                                switch_to_list(hwnd)
                                current_list_index += 1  # 更新當前列表索引

                    except Exception as e:
                        debug_print(f"處理檔案時發生錯誤: {str(e)}", color='light_red')
                        continue

                # 關閉剩餘的Chrome視窗
                if click_count > 0:
                    self.close_windows(click_count, initial_window_titles)

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
        self.hotkeys_enabled = False  # 改為 False，預設禁用
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
                debug_print("\n[DEBUG] ESC 按鍵被按下，停止執行", color='light_red', bold=True)
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
        self.stop_event.set()  # 設置停止事件即可，不需要等待線程結束

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

    @ensure_foreground_window
    def click_daily_report_tab(self, hwnd=None, window_title=None):
        """點擊每日報告標籤"""
        try:
            # 如果沒有提供視窗句柄，則獲取
            if hwnd is None or window_title is None:
                hwnd, window_title = self.selected_window or find_window_handle(Config.TARGET_WINDOW)[0]
            
            # 連接到視窗並找到標籤
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
            daily_report_tab = app.window(handle=hwnd).child_window(title="每日報告", control_type="TabItem")
            
            # 使用 calculate_center_position 計算中心點
            center_x, center_y = calculate_center_position(daily_report_tab.rectangle())
            if center_x is None or center_y is None:
                debug_print("無法計算標籤位置")
                return False
            
            click_at(center_x, center_y, clicks=2, interval=Config.SLEEP_INTERVAL, 
                    hwnd=hwnd, window_title=window_title)
            debug_print("已點擊每日報告標籤", color='light_yellow')
            return True
                
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
                debug_print(f"按下左鍵 {left_times} 次", color='light_yellow')
            for _ in range(left_times):
                pyautogui.press('left')
                time.sleep(Config.SLEEP_INTERVAL)
            
            if up_times > 0:
                debug_print(f"按下上鍵 {up_times} 次", color='light_yellow')
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
            ("點擊每日報告標籤", lambda: self.click_daily_report_tab(hwnd=hwnd, window_title=window_title)),
            ("設定字型大小", lambda: set_font_size()),
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title)),
            ("下載今日檔案，並收集列表", lambda: download_days_weeks(0, 0, '今日')),
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title)),
            ("下載昨日檔案，並收集列表", lambda: download_days_weeks(1, 0, '昨日')),
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title)),
            ("下載 1 週前檔案，並收集列表", lambda: download_days_weeks(0, 1, '1週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank(hwnd=hwnd, window_title=window_title)),
            ("下載 2 週前檔案，並收集列表", lambda: download_days_weeks(0, 1, '2週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank(hwnd=hwnd, window_title=window_title)),
            ("下載 4 週前檔案，並收集列表", lambda: download_days_weeks(0, 2, '4週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank(hwnd=hwnd, window_title=window_title)),
            ("下載 8 週前檔案，並收集列表", lambda: download_days_weeks(0, 4, '8週前')),
            ("點擊日歷空白處", lambda: start_click_calendar_blank(hwnd=hwnd, window_title=window_title)),
            ("鍵盤向下 X8", lambda: [pyautogui.press('down') or time.sleep(0.2) for _ in range(8)]),
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title)),
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
                debug_print("任務已停止", color='light_yellow')
                break
            
            debug_print(f"步驟{i}: {step_name}", color='light_yellow')
            step_func()
            time.sleep(1)
        
        debug_print("連續任務執行完成", color='light_green')
        self.stop_esc_listener()

    def download_current_list(self):
        """下載當前列表檔案"""
        self.should_stop = False
        self.stop_event.clear()
        reset_mouse_position()  # 重置滑鼠位置記錄
        
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
            debug_print("停止檢測列表刷新", color='light_yellow')
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
            
            # 預加載所有列表區域和檔案
            list_areas = {}
            list_files = {}
            list_types = {
                'morning': '晨會報告',
                'research': '研究報告',
                'industry': '產業報告'
            }
            
            # 一次性獲取所有列表的檔案
            for list_type, list_name in list_types.items():
                try:
                    list_areas[list_type] = main_window.child_window(
                        auto_id=f"listBox{list_type.capitalize()}Reports"
                    )
                    list_files[list_type] = get_list_items_by_id(main_window, list_type)
                except Exception as e:
                    debug_print(f"獲取 {list_name} 列表時發生錯誤: {str(e)}", color='light_red')
                    continue
            
            # 顯示結果
            for list_type, list_name in list_types.items():
                if list_type in list_files:
                    debug_print(f"\n=== {list_name} ===", color='light_cyan')
                    for i, item in enumerate(list_files[list_type], 1):
                        debug_print(f"{i}. {item.window_text()}", color='light_green')
                    
        except Exception as e:
            debug_print(f"列出報告清單時發生錯誤: {str(e)}", color='light_red')

    def collect_current_list(self, list_name, hwnd):
        """從指定窗口收集列表項目"""
        if list_name and hwnd:
            try:
                app = PywinautoApp(backend="uia").connect(handle=hwnd)
                main_window = app.window(handle=hwnd)
                
                # 預加載所有列表區域和檔案信息
                list_areas = {}
                list_files = {}
                list_types = [
                    ('morning', '晨會報告'),
                    ('research', '研究報告'),
                    ('industry', '產業報告')
                ]
                
                self.collected_lists[list_name] = []
                
                # 一次性獲取所有列表的檔案
                for list_type, list_display_name in list_types:
                    try:
                        list_areas[list_type] = main_window.child_window(
                            auto_id=f"listBox{list_type.capitalize()}Reports"
                        )
                        list_files[list_type] = get_list_items_by_id(main_window, list_type)
                        if list_files[list_type]:
                            self.collected_lists[list_name].extend(
                                [f.window_text() for f in list_files[list_type] if f.window_text()]
                            )
                    except Exception as e:
                        debug_print(f"收集 {list_display_name} 列表時發生錯誤: {str(e)}", color='light_red')
                        continue
                        
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
            ("點擊每日報告標籤", lambda: self.click_daily_report_tab(hwnd=hwnd, window_title=window_title)),
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title) or time.sleep(2)),
            ("收集今日列表", lambda: self.collect_current_list('今日', hwnd)),
            ("向左 1 天", lambda: pyautogui.press('left') or time.sleep(1)), # 向左 1 天，等待 1 秒
            ("收集昨日列表", lambda: self.collect_current_list('昨日', hwnd)),
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title) or time.sleep(2)),
            ("向上 1 週", lambda: pyautogui.press('up') or time.sleep(1)), # 向上 1 週，等待 1 秒
            ("收集 1 週前列表", lambda: self.collect_current_list('1週前', hwnd)),
            ("向上 1 週", lambda: [pyautogui.press('up') or time.sleep(1) for _ in range(1)]), # 向上 1 週，等待 1 秒
            ("收集 2 週前列表", lambda: self.collect_current_list('2週前', hwnd)),
            ("向上 2 週", lambda: [pyautogui.press('up') or time.sleep(1) for _ in range(2)]), # 向上 2 週，等待 1 秒
            ("收集 4 週前列表", lambda: self.collect_current_list('4週前', hwnd)),
            ("向上 4 週", lambda: [pyautogui.press('up') or time.sleep(1) for _ in range(4)]), # 向上 4 週，等待 1 秒
            ("收集 8 週前列表", lambda: self.collect_current_list('8週前', hwnd)),
            ("鍵盤向下 X8", lambda: [pyautogui.press('down') or time.sleep(1) for _ in range(8)]), # 鍵盤向下 X8，每次等待 1 秒
            ("點擊今日", lambda: start_calendar_checker(0, hwnd=hwnd, window_title=window_title)),
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
            debug_print(f"執行: {step_name}", color='light_yellow')
            step_func()
        
    def register_essential_hotkeys(self):
        """註冊必要的快捷鍵（F12開關和關閉程式）"""
        keyboard.add_hotkey('ctrl+shift+f12', self.toggle_hotkeys)
        keyboard.add_hotkey('ctrl+shift+q', lambda: True)

    def toggle_hotkeys(self):
        """切換快捷鍵啟用狀態"""
        keyboard.unhook_all()  # 先取消所有快捷鍵
        
        self.hotkeys_enabled = not self.hotkeys_enabled
        if self.hotkeys_enabled:
            self.register_hotkeys()  # 重新註冊所有快捷鍵
            debug_print("已啟用所有快捷鍵", color='light_green')
        else:
            self.register_essential_hotkeys()  # 只註冊必要的快捷鍵

    def register_hotkeys(self):
        """註冊所有快捷鍵"""
        keyboard.add_hotkey('ctrl+shift+e', self.execute_sequence)
        keyboard.add_hotkey('ctrl+shift+f', self.download_current_list)
        keyboard.add_hotkey('ctrl+shift+g', start_list_area_checker)
        keyboard.add_hotkey('ctrl+shift+b', set_font_size)
        keyboard.add_hotkey('ctrl+shift+t', self.toggle_refresh_check)
        keyboard.add_hotkey('ctrl+shift+r', list_all_controls)
        keyboard.add_hotkey('ctrl+shift+m', monitor_clicks)
        keyboard.add_hotkey('ctrl+shift+k', self.monitor_chrome)
        keyboard.add_hotkey('ctrl+shift+f8', test_terminal_support)
        keyboard.add_hotkey('ctrl+shift+f9', self.collect_and_analyze_lists)
        keyboard.add_hotkey('ctrl+shift+f10', self.list_all_reports)
        keyboard.add_hotkey('ctrl+shift+f11', self.copy_today_files)
        keyboard.add_hotkey('ctrl+shift+f12', self.toggle_hotkeys)  # 新增：切換快捷鍵的快捷鍵

    def run(self):
        try:
            debug_print("=== 研究報告自動下載程式 ===", color='light_yellow', bold=True)
            
            self.register_essential_hotkeys()  # 使用新方法註冊必要快捷鍵
            
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
            debug_print("按下 CTRL + SHIFT + F8   測試終端機支援", color='light_green')
            debug_print("按下 CTRL + SHIFT + F9   收集並分析所有列表", color='light_green')
            debug_print("按下 CTRL + SHIFT + F10  列出所有報告清單", color='light_green')
            debug_print("按下 CTRL + SHIFT + F11  複製今日所有新檔案", color='light_green')
            debug_print("按下 CTRL + SHIFT + F12  切換快捷鍵啟用狀態", color='light_green')
            debug_print("按下 ESC                 停止下載", color='light_yellow')
            debug_print("按下 CTRL+SHIFT+Q        關閉程式", color='light_red', bold=True)
            debug_print("==================", color='light_cyan')

            self.scheduler = Scheduler(self.execute_sequence)
            scheduler_thread = self.scheduler.init_scheduler() # 初始化排程器
            schedule_times = Config.get_schedule_times() # 獲取排程時間
            debug_print("==================", color='light_cyan')

            debug_print("已禁用所有快捷鍵（按下 CTRL + SHIFT + F12 切換）", color='light_yellow')
            
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
    main()  # 再執行主程式
