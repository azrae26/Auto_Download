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
from get_list_area import start_list_area_checker, set_stop
from utils import debug_print, find_window_handle
from scheduler import Scheduler
from font_size_setter import set_font_size

class Config:
    """配置類，集中管理所有配置參數"""
    RETRY_LIMIT = 8  # 向上翻頁次數
    SLEEP_INTERVAL = 0.1  # 等待時間
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
    def ensure_window_visible(hwnd, window_title):
        try:
            if win32gui.IsIconic(hwnd):
                debug_print(f"視窗 '{window_title}' 已最小化，正在還原...")
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(Config.SLEEP_INTERVAL)
            
            try:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(Config.SLEEP_INTERVAL)
                return True
            except Exception as e:
                debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")
                return False
        except Exception as e:
            debug_print(f"確保視窗可見時發生錯誤: {str(e)}")
            return False

    @staticmethod
    def is_process_running(process_name):
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == process_name.lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return False

class MouseController:
    """處理滑鼠相關操作"""
    @staticmethod
    def move_to_safe_position():
        screen_width, screen_height = pyautogui.size()
        pyautogui.moveTo(screen_width // 2, screen_height // 2)
        time.sleep(Config.SLEEP_INTERVAL)

    @staticmethod
    def is_position_safe(current_x, current_y, last_x, last_y):
        if last_x is None or last_y is None:
            return True
        offset_x = abs(current_x - last_x)
        offset_y = abs(current_y - last_y)
        return offset_x <= Config.MOUSE_MAX_OFFSET and offset_y <= Config.MOUSE_MAX_OFFSET

    @staticmethod
    def click_file(center_x, center_y, is_first_click=False):
        pyautogui.moveTo(center_x, center_y)
        time.sleep(Config.SLEEP_INTERVAL)
        pyautogui.doubleClick()
        if not is_first_click:
            time.sleep(Config.SLEEP_INTERVAL * 2)
        else:
            time.sleep(Config.SLEEP_INTERVAL * 5)

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
                click_x = (rect.left + rect.right) // 2
                click_y = rect.top + 10
                
                # 移動到位置並點擊
                pyautogui.moveTo(click_x, int(click_y))
                time.sleep(Config.SLEEP_INTERVAL * 2)
                pyautogui.click()
                time.sleep(Config.SLEEP_INTERVAL * 2)
                
                debug_print(f"已點擊下一個列表位置: x={click_x}, y={int(click_y)}")
                
            else:
                debug_print("警告: 找不到足夠的列表區域")
            
        except Exception as e:
            debug_print(f"切換列表時發生錯誤: {str(e)}")
        
        self.after_tab_switch = True

    def navigate_to_file(self, file, main_list_area, hwnd, file_name):
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
        try:
            file_rect = file.rectangle()
            file_center_y = (file_rect.top + file_rect.bottom) // 2
            return (file_center_y >= list_area.top and file_center_y <= list_area.bottom)
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

    def handle_refresh(self, files, i, last_click_x=None, last_click_y=None):
        new_file_count = len(files)
        
        if self.is_date_switching:
            self.current_file_count = new_file_count
            return False
        
        if self.current_file_count != 0 and new_file_count != self.current_file_count:
            debug_print(f"檢測到列表刷新: 檔案數量從 {self.current_file_count} 變為 {new_file_count}")
            self.last_known_position = i
            
            if last_click_x is not None and last_click_y is not None:
                MouseController.move_to_safe_position()
            
            time.sleep(Config.SLEEP_INTERVAL * 5)
            self.current_file_count = new_file_count
            return True
        
        if self.current_file_count == 0:
            self.current_file_count = new_file_count
        
        return False

    def close_windows(self, count):
        """關閉指定數量的視窗"""
        if count <= 0:
            return
        
        time.sleep(0.3)  # 等待視窗完全打開
        
        try:
            # 按下 CTRL
            keyboard.press('ctrl')
            time.sleep(0.2)  # 等待 0.2秒 確保 CTRL 被按下
            
            # 按指定次數的 W
            for _ in range(count):
                keyboard.press('w')
                time.sleep(0.1)  # 間隔 0.1秒
                keyboard.release('w')
                time.sleep(0.1)  # 間隔 0.1秒
            
            # 釋放 CTRL
            keyboard.release('ctrl')
            time.sleep(0.1)  # 等待 0.1秒 所有視窗關閉
            
        except Exception as e:
            debug_print(f"關閉視窗時發生錯誤: {str(e)}")
            # 確保 CTRL 鍵被釋放
            keyboard.release('ctrl')

    def process_files(self, app, hwnd, window_title, should_stop_callback):
        try:
            if should_stop_callback():
                debug_print("[DEBUG] 下載開始前檢測到停止信號")
                return

            if not WindowHandler.ensure_window_visible(hwnd, window_title):
                debug_print("錯誤: 無法確保視窗可見")
                return

            main_window = app.window(handle=hwnd)
            
            # 獲取所有列表區域
            list_areas = start_list_area_checker()
            if not list_areas or len(list_areas) < 3:
                debug_print("警告: 無法獲取完整的列表區域資訊")
                return

            # 獲取每個列表區域的檔案
            all_files = []
            for i, list_area in enumerate(['左側列表', '中間列表', '右側列表']):
                debug_print(f"獲取{list_area}檔案...")
                files = main_window.child_window(control_type="List", found_index=i).descendants(control_type="ListItem")
                all_files.extend([(file, i) for file in files if not file.window_text().endswith("_公司")])
                debug_print(f"{list_area}檔案數量: {len(files)}")

            if not all_files:
                debug_print("警告: 沒有找到可下載的檔案")
                return

            debug_print(f"總共找到 {len(all_files)} 個可下載檔案")
            
            # 下載檔案
            is_first_click = True
            click_count = 0
            downloaded_files = set()  # 記錄已下載的檔案

            for file, list_index in all_files:
                if should_stop_callback():
                    return

                try:
                    file_name = file.window_text()
                    if file_name in downloaded_files:
                        continue

                    debug_print(f"正在下載: {file_name}")
                    
                    # 檢查檔案是否在可視範圍內
                    list_area = list_areas[list_index]
                    if not self.is_file_visible(file, list_area):
                        debug_print(f"檔案 '{file_name}' 不在可視範圍內，嘗試調整位置")
                        if not self.scroll_to_file(file, list_area, hwnd):
                            debug_print(f"無法使檔案 '{file_name}' 進入可視範圍，跳過")
                            continue

                    # 執行點擊
                    rect = file.rectangle()
                    center_x = (rect.left + rect.right) // 2
                    center_y = (rect.top + rect.bottom) // 2
                    
                    MouseController.click_file(center_x, center_y, is_first_click)
                    
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
                    if self.is_last_file_in_list(file, list_index, all_files):
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
            files = main_window.child_window(control_type="List", found_index=i).descendants(control_type="ListItem")
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
                    files = main_window.child_window(control_type="List", found_index=i).descendants(control_type="ListItem")
                    for file in files:
                        if file.window_text() == file_name:
                            # 使用與主下載相同的邏輯
                            rect = file.rectangle()
                            center_x = (rect.left + rect.right) // 2
                            center_y = (rect.top + rect.bottom) // 2
                            
                            MouseController.click_file(center_x, center_y, False)
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
            
            # 找到目標檔案所在的列表
            lists = main_window.descendants(control_type="List")
            current_list = None
            for lst in lists:
                rect = lst.rectangle()
                if (rect.left == list_area.left and 
                    rect.top == list_area.top and 
                    rect.right == list_area.right and 
                    rect.bottom == list_area.bottom):
                    current_list = lst
                    break
            
            if not current_list:
                debug_print("找不到對應的列表")
                return False
            
            # 獲取當前列表的所有檔案
            current_files = current_list.children(control_type="ListItem")
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
                
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.2)
                
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

    def check_esc_key(self):
        """監聽 ESC 按鍵"""
        while True:
            if keyboard.is_pressed('esc'):
                self.should_stop = True
                debug_print("\n[DEBUG] ESC 按鍵被按下，停止執行")
                break
            time.sleep(0.1)

    def select_window(self, index):
        self.should_stop = False
        self.stop_event.clear()
        
        esc_thread = threading.Thread(target=self.check_esc_key, daemon=True)
        esc_thread.start()
        
        target_windows = find_window_handle(Config.TARGET_WINDOW)
        if 1 <= index <= len(target_windows):
            self.selected_window = target_windows[index - 1]
            debug_print(f"\n已選擇視窗: {self.selected_window[1]}")
            hwnd, window_title = self.selected_window
            app = PywinautoApp(backend="uia").connect(handle=hwnd)
            self.file_processor.process_files(app, hwnd, window_title, lambda: self.should_stop)
        else:
            debug_print("無效的選擇")

    def is_weekday_2_to_5(self):
        weekday = datetime.now().weekday()
        return 1 <= weekday <= 4

    def execute_sequence(self):
        """執行連續任務"""
        self.should_stop = False
        
        # 啟動 ESC 監聽線程
        esc_thread = threading.Thread(target=self.check_esc_key, daemon=True)
        esc_thread.start()
        
        debug_print("開始執行連續任務...")
        
        # 每個步驟前都檢查是否要停止
        steps = [
            ("設定字型大小", lambda: set_font_size()),
            ("點擊今日日期", lambda: start_calendar_checker()),
            ("下載檔案", lambda: self.select_window(1)),
            ("再次點擊今日", lambda: start_calendar_checker()),
        ]
        
        # 如果是週二到週五，添加額外步驟
        if self.is_weekday_2_to_5():
            extra_steps = [
                ("按左鍵", lambda: pyautogui.press('left')),
                ("下載檔案", lambda: self.select_window(1)),
                ("點擊今日", lambda: start_calendar_checker()),
                ("按上鍵", lambda: pyautogui.press('up')),
                ("下載檔案", lambda: self.select_window(1))
            ]
            steps.extend(extra_steps)
        
        # 執行所有步驟
        for i, (step_name, step_func) in enumerate(steps, 1):
            if self.should_stop:
                debug_print("任務已停止")
                return
            
            debug_print(f"步驟{i}: {step_name}")
            step_func()
            time.sleep(1)  # 每個步驟之間暫停1秒
        
        debug_print("連續任務執行完成")

    def download_current_list(self):
        self.should_stop = False
        self.stop_event.clear()
        
        debug_print("開始下載當前列表檔案...")
        
        esc_thread = threading.Thread(target=self.check_esc_key, daemon=True)
        esc_thread.start()
        
        MouseController.move_to_safe_position()
        self.select_window(1)

    def run(self):
        try:
            debug_print("=== 自動下載程式 ===")
            debug_print("按下 CTRL+Q 或 CTRL+E 開始連續下載任務")
            debug_print("按下 CTRL+D 下載當前列表檔案")
            debug_print("按下 CTRL+G 檢測檔案列表區域")
            debug_print("按下 ESC 停止下載")
            debug_print(f"請確保 {Config.PROCESS_NAME} 已開啟且視窗可見")
            
            keyboard.add_hotkey('ctrl+q', self.execute_sequence)
            keyboard.add_hotkey('ctrl+e', self.execute_sequence)
            keyboard.add_hotkey('ctrl+d', self.download_current_list)
            keyboard.add_hotkey('ctrl+g', start_list_area_checker)
            keyboard.add_hotkey('ctrl+b', set_font_size)
            
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
