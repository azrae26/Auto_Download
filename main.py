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
from calendar_checker import start_calendar_checker
from get_list_area import start_list_area_checker, set_stop, list_all_controls, monitor_clicks
from utils import debug_print, find_window_handle, ensure_foreground_window, get_list_items, calculate_center_position, refresh_checking, start_refresh_check, stop_refresh_check, click_at, move_to_safe_position, check_mouse_movement
from scheduler import Scheduler
from font_size_setter import set_font_size

class Config:
    """配置類，集中管理所有配置參數"""
    RETRY_LIMIT = 8  # 向上翻頁次數
    SLEEP_INTERVAL = 0.1  # 基本等待時間為 0.1 秒
    CLICK_BATCH_SIZE = 5  # 批次點擊次數
    MOUSE_MAX_OFFSET = 100  # 滑鼠最大偏移量
    TARGET_WINDOW = "stocks"
    PROCESS_NAME = "DOstocksBiz.exe"

    @staticmethod
    def get_schedule_times():
        return ["10:00"] # 排程時間

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

    def switch_to_next_list(self, hwnd):
        """切換到下一個列表"""
        debug_print("開始切換列表: 點擊左鍵")
        
        try:
            # 連接到視窗
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
            main_window = app.window(handle=hwnd)
            
            # 獲取當前列表區域
            lists = main_window.descendants(control_type="List")
            if len(lists) >= 2:  # 確保至少有兩個列表
                # 獲取下一個列表的位置
                next_list = lists[1]  # 從左側列表切換到中間列表
                rect = next_list.rectangle()
                
                # 計算點擊位置（列表頂部往下 10px 的位置）
                center_x, center_y = calculate_center_position(rect)
                if center_x is None or center_y is None:
                    debug_print("無法計算列表位置")
                    return
                click_y = rect.top + 10
                
                # 移動到位置並點擊
                click_at(center_x, click_y, clicks=1, hwnd=hwnd, window_title=win32gui.GetWindowText(hwnd))
                time.sleep(Config.SLEEP_INTERVAL * 2)
                
                debug_print(f"已點擊下一個列表位置: x={center_x}, y={click_y}")
                
            else:
                debug_print("警告: 找不到足夠的列表區域")
            
        except Exception as e:
            debug_print(f"切換列表時發生錯誤: {str(e)}")
        
        self.after_tab_switch = True

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
                    self.switch_to_next_list(hwnd)
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
                    self.switch_to_next_list(hwnd)
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
    def is_file_visible(file, list_area):
        """檢查檔案是否在可視範圍內"""
        try:
            file_rect = file.rectangle() # 獲取檔案的矩形
            list_rect = list_area.rectangle()  # 獲取列表區域的矩形
            # file_rect.top >= list_rect.top 檢查檔案頂部是否在列表區域的上方
            # file_rect.bottom <= list_rect.bottom 檢查檔案底部是否在列表區域的下方
            return (file_rect.top >= list_rect.top and file_rect.bottom <= list_rect.bottom) # 檢查檔案是否在列表區域的可視範圍內
        except Exception as e:
            debug_print(f"檢查檔案可見性時發生錯誤: {str(e)}")
            return False

    @staticmethod
    def is_last_file_in_current_list(file, current_list_items):
        """判斷是否為當前列表的最後一個檔案"""
        try:
            # 獲取當前檔案在列表中的索引
            current_index = current_list_items.index(file)
            # 如果是列表中的最後一個項目
            return current_index == len(current_list_items) - 1
        except Exception as e:
            debug_print(f"檢查檔案位置時發生錯誤: {str(e)}")
            return False

    def close_windows(self, count):
        """關閉指定數量的視窗"""
        if count <= 0:
            return
        
        time.sleep(0.3)  # 等待視窗完全打開
        
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
            debug_print(f"關閉視窗時發生錯誤: {str(e)}")
            # 確保 CTRL 鍵被釋放
            keyboard.release('ctrl')

    def process_files(self, app, hwnd, window_title, should_stop_callback):
        """處理檔案"""
        try:
            if should_stop_callback():
                debug_print("[DEBUG] 下載開始前檢測到停止信號")
                return

            if not ensure_foreground_window(hwnd, window_title):
                debug_print("錯誤: 無法確保視窗可見")
                return

            main_window = app.window(handle=hwnd)
            
            # 獲取所有列表區域
            list_areas = start_list_area_checker()
            if not list_areas:
                debug_print("警告: 無法獲取列表區域資訊")
                return

            # 獲取每個列表區域的檔案
            all_files = []
            valid_areas = []  # 儲存有效的列表區域
            
            for i, list_area in enumerate(['左側列表', '中間列表', '右側列表']):
                debug_print(f"獲取{list_area}檔案...")
                files = get_list_items(main_window, i)
                if files:  # 只處理有檔案的列表
                    file_count = len([f for f in files if not f.window_text().endswith("_公司")])
                    if file_count > 0:  # 確保有可下載的檔案
                        valid_files = [
                            (file, len(valid_areas)) for file in files if not file.window_text().endswith("_公司")]
                        all_files.extend(valid_files)
                        valid_areas.append(list_areas[i])  # 儲存有效的列表區域
                        debug_print(f"{list_area}案數量: {file_count}")
                    else:
                        debug_print(f"{list_area}沒有可下載的檔案")
                else:
                    debug_print(f"{list_area}是空白的，跳過")

            if not all_files:
                debug_print("警告: 沒有找到可下載的檔案")
                return

            debug_print(f"總共找到 {len(all_files)} 個可下載檔案")
            
            # 下載檔案
            is_first_click = True
            click_count = 0
            downloaded_files = set()

            for file, area_index in all_files:
                if should_stop_callback():
                    return

                try:
                    file_name = file.window_text()
                    if file_name in downloaded_files:
                        continue

                    debug_print(f"正在下載: {file_name}")
                    
                    # 使用對應的有效列表區域
                    list_area = valid_areas[area_index]
                    if not self.is_file_visible(file, list_area):
                        debug_print(f"檔案 '{file_name}' 不在可視範圍內，嘗試調整位置")
                        if not self.scroll_to_file(file, list_area, hwnd):
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
                        target_element=file  # 傳入目標檔案元素
                    )
                    
                    # 如果按下 CTRL+B 設定字型大小，則不計算點擊次數
                    if not is_first_click:
                        click_count += 1
                        if click_count == Config.CLICK_BATCH_SIZE:
                            self.close_windows(Config.CLICK_BATCH_SIZE)
                            click_count = 0
                    else:
                        is_first_click = False
                        time.sleep(Config.SLEEP_INTERVAL * 5)

                    downloaded_files.add(file_name)
                    
                    # 如果是當前列表的最後一個檔案，切換到下一個列表
                    if self.is_last_file_in_list(file, area_index, all_files):
                        debug_print(f"切換到下一個列表")
                        self.navigator.switch_to_next_list(hwnd)  # 使用 navigator 來切換列表

                except Exception as e:
                    debug_print(f"處理檔案時發生錯誤: {str(e)}")
                    continue

            # 關閉剩餘視窗
            if click_count > 0:
                self.close_windows(click_count)

            # 檢查是否有漏掉的檔案
            debug_print("檢查是否有漏掉的檔案...")
            new_files = set(file.window_text() for file, _ in self.get_all_files(main_window))
            missed_files = new_files - downloaded_files
            
            if missed_files:
                debug_print(f"發現 {len(missed_files)} 個漏掉的檔案，開始下載...")
                # 下載漏掉的檔案
                self.download_missed_files(missed_files, main_window, hwnd)

            debug_print("所有檔案下載完成")

        except Exception as e:
            debug_print(f"發生錯誤: {str(e)}")

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
            files = get_list_items(main_window, i)
            all_files.extend([(file, i) for file in files if not file.window_text().endswith("_公司")])
        return all_files

    def download_missed_files(self, missed_files, main_window, hwnd):
        """下載漏掉的檔案"""
        click_count = 0
        for file_name in missed_files:
            if self.should_stop:
                return
                
            try:
                # 在所有列表中尋找檔案
                for i in range(3):
                    files = get_list_items(main_window, i)
                    for file in files:
                        if file.window_text() == file_name:
                            # 使用與主下載相同的邏輯
                            rect = file.rectangle()
                            center_x, center_y = calculate_center_position(rect)
                            if center_x is None or center_y is None:
                                debug_print("無法計算檔案位置，跳過此檔案")
                                continue
                            click_at(
                                x=center_x, 
                                y=center_y, 
                                is_first_click=False,
                                clicks=2,
                                interval=Config.SLEEP_INTERVAL,
                                hwnd=hwnd,
                                window_title=win32gui.GetWindowText(hwnd)
                            )
                            click_count += 1
                            
                            if click_count == Config.CLICK_BATCH_SIZE:
                                self.close_windows(Config.CLICK_BATCH_SIZE)
                                click_count = 0
                            
                            break
            except Exception as e:
                debug_print(f"下載漏掉的檔案時發生錯誤: {str(e)}")
                continue
        
        # 關閉剩餘視窗
        if click_count > 0:
            self.close_windows(click_count)

    def scroll_to_file(self, file, list_area, hwnd):
        """滾動直到檔案進入可視範圍"""
        try:
            # 連接到視窗
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
            main_window = app.window(handle=hwnd)
            
            # 獲取列表區域的矩形
            list_rect = list_area.rectangle()
            
            # 獲取當前列表的所有檔案
            current_files = get_list_items(list_area)
            target_name = file.window_text()
            
            # 找到目標檔案的索引
            target_index = -1
            for i, f in enumerate(current_files):
                if f.window_text() == target_name:
                    target_index = i
                    break
            
            if target_index == -1:
                debug_print(f"在列表中找不到檔案: {target_name}")
                return False
            
            # 找到當前可見的第一個檔案的索引
            visible_index = -1
            for i, f in enumerate(current_files):
                if self.is_file_visible(f, list_area):
                    visible_index = i
                    debug_print(f"當前可見的第一個檔案索引: {i}")
                    debug_print(f"檔案名稱: {f.window_text()}")
                    break
            
            if visible_index == -1:
                debug_print("找不到可見的檔案")
                return False
            
            debug_print(f"目標檔案索引: {target_index}")
            debug_print(f"可見檔案索引: {visible_index}")
            
            # 根據索引決定滾動方向
            max_attempts = 8
            for attempt in range(max_attempts):
                if self.is_file_visible(file, list_area):
                    debug_print("檔案已在可視範圍內")
                    return True
                
                if not ensure_foreground_window(hwnd):
                    debug_print("警告: 無法確保視窗在前景")
                    
                if target_index < visible_index:
                    debug_print(f"目標檔案在可見區域上方，向上翻頁 (第 {attempt + 1} 次)")
                    pyautogui.press('pageup')
                else:
                    debug_print(f"目標檔案在可見區域下方，向下翻頁 (第 {attempt + 1} 次)")
                    pyautogui.press('pagedown')
                time.sleep(0.5)
                
                # 更新可見檔案的索引
                for i, f in enumerate(current_files):
                    if self.is_file_visible(f, list_area):
                        visible_index = i
                        break
            
            return False
                
        except Exception as e:
            debug_print(f"滾動到檔案位置時發生錯誤: {str(e)}")
            return False

class MainApp:
    """主應用程序類"""
    def __init__(self):
        self.file_processor = FileProcessor()
        self.selected_window = None
        self.scheduler = None
        self.should_stop = False
        self.stop_event = threading.Event()
        self.esc_thread = None  # 新增：保存 ESC 監聽線程的引用

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
        """選擇視窗"""
        self.start_esc_listener()  # 使用新的啟動方法
        
        target_windows = find_window_handle(Config.TARGET_WINDOW)
        if 1 <= index <= len(target_windows):
            self.selected_window = target_windows[index - 1]
            debug_print(f"\n已選擇視窗: {self.selected_window[1]}")
            hwnd, window_title = self.selected_window
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
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
                debug_print("警告: 無法確保視窗前景")
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
                debug_print("已點擊每日報告標籤")
                return True
            else:
                debug_print("無法獲取每日報告標籤位置")
                return False
                
        except Exception as e:
            debug_print(f"點擊每日報告標籤時發生錯誤: {str(e)}")
            return False

    def execute_sequence(self):
        """執行連續任務"""
        self.start_esc_listener()
        
        debug_print("開始執行連續任務...")
        
        # 先獲取視窗句柄
        hwnd, window_title = self.selected_window or find_window_handle(Config.TARGET_WINDOW)[0]
        
        steps = [
            ("點擊每日報告標籤", lambda: self.click_daily_report_tab(hwnd, window_title)),  # 傳入已獲取的句柄
            ("設定字型大小", lambda: set_font_size()),
            ("點擊今日日期", lambda: start_calendar_checker()),
            ("下載檔案", lambda: self.select_window(1)),
            ("再次點擊今日", lambda: start_calendar_checker()),
            ("按左鍵", lambda: pyautogui.press('left')),
            ("下載檔案", lambda: self.select_window(1)),
            ("點擊今日", lambda: start_calendar_checker()),
            ("按上鍵", lambda: pyautogui.press('up')),
            ("下載檔案", lambda: self.select_window(1))
        ]

        for i, (step_name, step_func) in enumerate(steps, 1):
            if self.should_stop:
                debug_print("任務已停止")
                break
            
            debug_print(f"步驟{i}: {step_name}")
            step_func()
            time.sleep(1)
        
        debug_print("連續任務執行完成")
        self.stop_esc_listener()  # 操作完成後停止監聽

    def download_current_list(self):
        """下載當前列表檔案"""
        self.should_stop = False
        self.stop_event.clear()
        
        debug_print("開始載當前列表檔案...")
        
        self.start_esc_listener()  # 使用新的啟動方法
        
        move_to_safe_position()
        self.select_window(1)

    def toggle_refresh_check(self):
        """切換列表刷新檢測"""
        if not refresh_checking:
            hwnd = self.selected_window[0] if self.selected_window else None
            threading.Thread(
                target=start_refresh_check,
                args=(hwnd,),
                daemon=True
            ).start()
        else:
            stop_refresh_check()

    def run(self):
        try:
            debug_print("=== 自動下載程式 ===")
            """""
            keyboard.add_hotkey('ctrl+e', self.execute_sequence)
            keyboard.add_hotkey('ctrl+d', self.download_current_list)
            keyboard.add_hotkey('ctrl+g', start_list_area_checker)
            keyboard.add_hotkey('ctrl+b', set_font_size)
            keyboard.add_hotkey('ctrl+t', self.toggle_refresh_check)
            keyboard.add_hotkey('ctrl+r', list_all_controls)
            keyboard.add_hotkey('ctrl+m', monitor_clicks)
            debug_print("按下 CTRL+E 開始連續下載任務")
            debug_print("按下 CTRL+D 下載當前列表檔案")
            debug_print("按下 CTRL+G 檢測檔案列表區域")
            debug_print("按下 CTRL+B 設定字型大小")
            debug_print("按下 CTRL+T 切換列表刷新檢測")
            debug_print("按下 CTRL+R 列出所有控件")
            debug_print("按下 CTRL+M 開始監控滑鼠點擊")
            """""
            debug_print("按下 ESC 停止下載")
            self.scheduler = Scheduler(self.execute_sequence)
            scheduler_thread = self.scheduler.init_scheduler()
            schedule_times = Config.get_schedule_times()
            
            keyboard.wait('ctrl+c')
                
        except KeyboardInterrupt:
            debug_print("\n程式已結束")
        finally:
            keyboard.unhook_all()

if __name__ == "__main__":
    try:
        app = MainApp()
        app.run()
    except KeyboardInterrupt:
        debug_print("\n程式已結束")
    except Exception as e:
        debug_print(f"程式發生未預期的錯誤: {str(e)}")
