import keyboard
import pywinauto
import pyautogui
import time
from pywinauto.application import Application
import sys
import psutil
import win32gui
import win32con
import threading
from queue import Queue
from datetime import datetime, timedelta
from calendar_checker import start_calendar_checker
from get_list_area import start_list_area_checker
from utils import debug_print, debug_queue
from scheduler import Scheduler

# 全域變數
should_stop = False
stop_event = threading.Event()
last_refresh_time = None
REFRESH_COOLDOWN = 2  # 刷新冷卻時間（秒）
current_file_count = 0
last_known_position = 0
is_date_switching = False  # 用於標記是否正在切換日期

pyautogui.FAILSAFE = False

def check_esc_key():
    global should_stop
    while True:
        if keyboard.is_pressed('esc'):
            should_stop = True
            stop_event.set()
            debug_print("\n[DEBUG] ESC 按鍵被按下，設置停止標誌")
            break
        time.sleep(0.1)

def is_process_running(process_name):
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'].lower() == process_name.lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

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

selected_window = None

def select_window(index):
    global selected_window, should_stop, stop_event
    should_stop = False
    stop_event.clear()
    
    # 啟動按鍵監聽執行緒
    esc_thread = threading.Thread(target=check_esc_key, daemon=True)
    esc_thread.start()
    
    target_windows = find_window_handle("stocks")
    if 1 <= index <= len(target_windows):
        selected_window = target_windows[index - 1]
        debug_print(f"\n已選擇視窗: {selected_window[1]}")
        download_files()
    else:
        debug_print("無效的選擇")

def safe_sleep(seconds):
    """分段睡眠，同時檢查停止標誌"""
    interval = 0.1
    for _ in range(int((seconds/2) / interval)):  # 將等待時間減半
        if should_stop:
            debug_print("[DEBUG] 在等待期間檢測停止信號")
            return False
        time.sleep(interval)
    return True

def ensure_window_visible(hwnd, window_title):
    """確保視窗可見且在前景"""
    try:
        if win32gui.IsIconic(hwnd):
            debug_print(f"視窗 '{window_title}' 已最小化，正在還原...")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)  # 從0.5改為0.2
        
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.2)  # 從0.5改為0.2
            return True
        except Exception as e:
            debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")
            return False
    except Exception as e:
        debug_print(f"確保視窗可見時發生錯誤: {str(e)}")
        return False

def handle_refresh(files, i, last_click_x=None, last_click_y=None):
    """處理列表刷新的情況"""
    global last_refresh_time, current_file_count, last_known_position, is_date_switching
    
    current_time = datetime.now()
    new_file_count = len(files)
    
    # 如果是日期切換造成的檔案數量變化，直接返回 False
    if is_date_switching:
        current_file_count = new_file_count
        last_refresh_time = current_time
        return False
    
    # 只在檔案數量突然改變時視為刷新
    if current_file_count != 0 and new_file_count != current_file_count:
        debug_print(f"檢測到列表刷新: 檔案數量從 {current_file_count} 變為 {new_file_count}")
        
        # 保存當前進度
        last_known_position = i
        
        # 如果有上次的點擊位置，先將滑鼠移回安全位置
        if last_click_x is not None and last_click_y is not None:
            screen_width, screen_height = pyautogui.size()
            pyautogui.moveTo(screen_width // 2, screen_height // 2)
            time.sleep(0.2)
        
        time.sleep(0.5)  # 等待列表穩定
        current_file_count = new_file_count
        last_refresh_time = current_time
        return True
    
    # 如果沒有檢測到刷新，只更新檔案數量
    if current_file_count == 0:
        current_file_count = new_file_count
    
    return False

def is_file_visible(file, list_area):
    """檢查檔案是否在對應列表的可視範圍內"""
    try:
        # 獲取檔案的位置
        file_rect = file.rectangle()
        file_center_y = (file_rect.top + file_rect.bottom) // 2
        
        # 檢查檔案中心點是否在列表可視範圍內
        is_visible = (file_center_y >= list_area.top and 
                     file_center_y <= list_area.bottom)
        
        return is_visible
    except Exception as e:
        debug_print(f"檢查檔案可見性時發生錯誤: {str(e)}")
        return False

def is_last_file_in_current_list(file_name):
    """檢查是否為當前列表的最後一個檔案"""
    # 晨訊列表的最後一個檔案通常是"群益"結尾
    if file_name.startswith("晨訊") and file_name.endswith("群益"):
        return True
    return False

def switch_to_next_list(hwnd):
    """切換到下一個列表並設置向上搜尋"""
    debug_print("開始切換列表: 點擊左鍵")
    pyautogui.click()
    time.sleep(0.5)
    
    debug_print("按下 3 次 TAB 切換列表")
    for i in range(3):
        pyautogui.press('tab')
        time.sleep(0.2)
    time.sleep(0.5)
    
    # 直接設置為向上搜尋模式
    global searching_up, down_retry_count, up_retry_count, after_tab_switch
    searching_up = True
    down_retry_count = 8  # 跳過向下搜尋
    up_retry_count = 0
    after_tab_switch = True  # 新增標記，表示剛按完 TAB

def reset_search_state():
    """重置搜尋狀態為向下搜尋"""
    global searching_up, down_retry_count, up_retry_count
    searching_up = False
    down_retry_count = 0
    up_retry_count = 0

def download_files():
    try:
        global selected_window, should_stop, last_known_position
        
        if should_stop:
            debug_print("[DEBUG] 下載開始前檢測到停止信號")
            return

        # 檢查程式是否運行
        if not is_process_running("DOstocksBiz.exe"):
            debug_print("錯誤: DOstocksBiz.exe 未行，請先開啟程式")
            return

        # 如果還沒有選擇視窗，顯示可用的視窗列表
        if not selected_window:
            target_windows = find_window_handle("stocks")
            
            if not target_windows:
                debug_print("錯誤: 找不到相關視窗")
                return

            debug_print("\n找到以下視窗:")
            for i, (_, title) in enumerate(target_windows, 1):
                debug_print(f"{i}. {title}")
            debug_print("\n請按數字鍵 1-{} 選擇正確的視窗".format(len(target_windows)))
            return

        hwnd, window_title = selected_window

        if should_stop:
            debug_print("[DEBUG] 視窗選擇後檢測到停止信號")
            return

        # 確保視窗可見
        if not ensure_window_visible(hwnd, window_title):
            debug_print("錯誤: 無法確保視窗可見")
            return

        if should_stop:
            debug_print("[DEBUG] 視窗前景化後檢測到停止信號")
            return

        # 連接到程式
        app = Application(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        
        # 獲取所有檔案元素
        debug_print("正在掃描檔案列表...")
        files = main_window.descendants(control_type="ListItem")
        
        if not files:
            debug_print("警告: 沒有找到可下載的檔案")
            return

        debug_print(f"找到 {len(files)} 個檔案")
        
        # 記錄上一次成功的點擊位置
        last_click_x = None
        last_click_y = None
        # 設定允許的最大偏移距離（像素）
        MAX_MOUSE_OFFSET = 100
        
        # 獲取列表區域的位置資訊
        list_areas = start_list_area_checker()
        if not list_areas or len(list_areas) < 3:
            debug_print("警告: 無法獲取完整的列表區域資訊")
            return

        # 使用第二個列表區域（主要的檔案列表）
        main_list_area = list_areas[1]  # 第二個列表區域
        
        # 添加一個標記，用於識別是否為第一次點擊
        is_first_click = True
        # 添加計數器，用於追蹤連續點擊次數
        click_count = 0
        
        # 將搜尋狀態變數移到函數開頭
        global searching_up, down_retry_count, up_retry_count, after_tab_switch
        searching_up = False
        down_retry_count = 0
        up_retry_count = 0
        after_tab_switch = False  # 初始化 TAB 切換標記
        
        # 遍歷並下載所有檔案
        i = last_known_position
        while i < len(files):
            if should_stop:
                debug_print("[DEBUG] 檔案迴圈中檢測到止信號")
                return
            
            try:
                # 處理刷新，傳入最後的點擊位置
                if handle_refresh(files, i, last_click_x, last_click_y):
                    # 重新獲取檔案列表
                    files = main_window.descendants(control_type="ListItem")
                    
                    # 如果位置超出新的列表範圍，重置到最後一個有效位置
                    if last_known_position >= len(files):
                        last_known_position = max(0, len(files) - 1)
                    
                    # 恢復到上次的位置
                    if last_known_position > 0:
                        debug_print(f"嘗試恢復到位置 {last_known_position}")
                        # 使用 Page Down 快速到達目標位置附近
                        pages_needed = last_known_position // 10
                        for _ in range(pages_needed):
                            pyautogui.press('pagedown')
                            time.sleep(0.1)
                        
                        # 重新獲取當前檔案的位置
                        try:
                            current_file = files[i]
                            rect = current_file.rectangle()
                            center_x = (rect.left + rect.right) // 2
                            center_y = (rect.top + rect.bottom) // 2
                            
                            # 移動到檔案位置
                            pyautogui.moveTo(center_x, center_y)
                            time.sleep(0.2)
                        except Exception as e:
                            debug_print(f"恢復滑鼠位置時發生錯誤: {str(e)}")
                    
                    continue
                
                # 每次點擊前確保視窗可見
                if not ensure_window_visible(hwnd, window_title):
                    debug_print("警告: 無法確保視窗可見，重試當前檔案")
                    continue
                
                file = files[i]
                file_name = file.window_text()
                
                # 重置搜尋狀態（除非是剛按完 TAB）
                if not after_tab_switch:
                    reset_search_state()
                
                while not is_file_visible(file, main_list_area):
                    if should_stop:
                        return
                    
                    # 根據是否剛按完 TAB 決定搜尋方向
                    if after_tab_switch:
                        # TAB 切換後直接向上搜尋
                        if up_retry_count < 8:
                            debug_print(f"檔案 '{file_name}' 不在可視範圍內，向上翻頁 (第 {up_retry_count + 1} 次)")
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.2)
                            pyautogui.press('pageup')
                            time.sleep(0.5)
                            up_retry_count += 1
                        else:
                            # 向上找完都找不到，切換到下一個列表
                            debug_print(f"無法在當前列表找到檔案 '{file_name}'，嘗試切換到下一個列表")
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.2)
                            switch_to_next_list(hwnd)
                            continue
                    else:
                        # 一般情況：先向下找再向上找
                        if not searching_up and down_retry_count < 8:
                            debug_print(f"檔案 '{file_name}' 不在可視範圍內，向下翻頁 (第 {down_retry_count + 1} 次)")
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.2)
                            pyautogui.press('pagedown')
                            time.sleep(0.5)
                            down_retry_count += 1
                        elif not searching_up and down_retry_count >= 8:
                            debug_print("向下找不到，開始向上翻頁尋找")
                            searching_up = True
                        elif searching_up and up_retry_count < 8:
                            debug_print(f"檔案 '{file_name}' 不在可視範圍內，向上翻頁 (第 {up_retry_count + 1} 次)")
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.2)
                            pyautogui.press('pageup')
                            time.sleep(0.5)
                            up_retry_count += 1
                        else:
                            # 都找不到，切換到下一個列表
                            debug_print(f"無法在當前列表找到檔案 '{file_name}'，嘗試切換到下一個列表")
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.2)
                            switch_to_next_list(hwnd)
                            continue
                
                # 找到檔案後，重置 TAB 切換標記
                after_tab_switch = False
                
                # 檢查檔名是否以"_公司"結尾
                if file_name.endswith("_公司"):
                    debug_print(f"跳過檔案 ({i+1}/{len(files)}): {file_name} (檔名以_公司結尾)")
                    i += 1
                    continue
                
                # 獲取檔案位置
                rect = file.rectangle()
                center_x = (rect.left + rect.right) // 2
                center_y = (rect.top + rect.bottom) // 2
                
                # 檢查滑鼠位置是否偏離太多
                current_mouse_x, current_mouse_y = pyautogui.position()
                if last_click_x is not None and last_click_y is not None:
                    offset_x = abs(current_mouse_x - last_click_x)
                    offset_y = abs(current_mouse_y - last_click_y)
                    
                    if offset_x > MAX_MOUSE_OFFSET or offset_y > MAX_MOUSE_OFFSET:
                        debug_print(f"檢測到滑鼠偏移過大 (x:{offset_x}, y:{offset_y})")
                        debug_print("暫停下載...")
                        time.sleep(0.5)
                        
                        debug_print(f"重新定位到當前檔案: {file_name}")
                        
                        # 將滑鼠移到螢幕中央
                        screen_width, screen_height = pyautogui.size()
                        pyautogui.moveTo(screen_width // 2, screen_height // 2)
                        time.sleep(0.2)
                        
                        # 將滑鼠移到目標檔案位置
                        pyautogui.moveTo(center_x, center_y)
                        time.sleep(0.2)
                        
                        debug_print("繼續下載...")
                        continue
                
                debug_print(f"正在下載 ({i+1}/{len(files)}): {file_name}")
                
                if should_stop:
                    debug_print("[DEBUG] 點擊前檢測到停止信號")
                    return
                
                # 先移動到目標位置
                pyautogui.moveTo(center_x, center_y)
                time.sleep(0.2)
                
                # 執行點擊並記錄位置
                pyautogui.doubleClick()
                last_click_x = center_x
                last_click_y = center_y
                
                # 在雙擊後執行 CTRL+W，但跳過第一次
                if not is_first_click:
                    click_count += 1
                    if click_count == 5:
                        time.sleep(0.3)
                        for _ in range(5):
                            pyautogui.hotkey('ctrl', 'w')
                            time.sleep(0.2)
                        click_count = 0
                else:
                    is_first_click = False
                    time.sleep(0.5)

                if not safe_sleep(0.1):
                    return
                
                # 如果是最後一個檔案，下載完後再切換列表
                if is_last_file_in_current_list(file_name):
                    debug_print(f"檔案 '{file_name}' 是當前列表的最後一個檔案，準備切換到下一個列表")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                    switch_to_next_list(hwnd)  # 切換列表並自動設置搜尋狀態
                
                i += 1
                last_known_position = i
                
            except Exception as e:
                debug_print(f"下載檔案時發生錯誤: {str(e)}")
                i += 1
                continue
            
        # 處理最後剩餘的未關閉視窗
        if click_count > 0:
            time.sleep(0.3)
            for _ in range(click_count):
                pyautogui.hotkey('ctrl', 'w')
                time.sleep(0.1)
        
        debug_print("所有檔案下載完成")
        last_known_position = 0  # 重置位置
        
    except Exception as e:
        debug_print(f"發生錯誤: {str(e)}")
        debug_print("\n請確保:")
        debug_print("1. DOstocksBiz.exe 已經開啟")
        debug_print("2. 視窗未最小化")
        debug_print("3. 已選擇正確的視窗")
    finally:
        if should_stop:
            debug_print("[DEBUG] 下載函數結束時的最終停止信號檢查")

def is_weekday_2_to_5():
    weekday = datetime.now().weekday()  # 0是週一，6是週日
    return 1 <= weekday <= 4  # 週二到週五

def execute_sequence():
    global should_stop, stop_event, is_date_switching
    
    # 重置停止標誌
    should_stop = False
    stop_event.clear()
    
    debug_print("開始執行連續任務...")
    
    # 啟動按鍵監聽執行緒
    esc_thread = threading.Thread(target=check_esc_key, daemon=True)
    esc_thread.start()
    
    # 將滑鼠移到螢幕中央的函數
    def move_mouse_to_safe_position():
        screen_width, screen_height = pyautogui.size()
        pyautogui.moveTo(screen_width // 2, screen_height // 2)
        time.sleep(0.5)
    
    # 1. 點擊今日日期
    debug_print("步驟1: 點擊今日日期")
    is_date_switching = True  # 設置日期切換標記
    move_mouse_to_safe_position()
    start_calendar_checker()
    if not safe_sleep(1):
        return
    is_date_switching = False  # 重置標記
        
    # 2. 下載檔案
    debug_print("步驟2: 下載檔案")
    move_mouse_to_safe_position()
    select_window(1)
    if not safe_sleep(1):
        return
        
    # 3. 再次點擊今日
    debug_print("步驟3: 再次點擊今日")
    is_date_switching = True  # 設置日期切換標記
    move_mouse_to_safe_position()
    start_calendar_checker()
    if not safe_sleep(1):
        return
    is_date_switching = False  # 重置標記
    
    # 4. 如果今天是週二到週五，執行額外步驟
    if is_weekday_2_to_5():
        debug_print("步驟4: 執行週二到週五的額外步驟")
        
        move_mouse_to_safe_position()
        pyautogui.press('left')
        if not safe_sleep(1):
            return
        select_window(1)
        if not safe_sleep(1):
            return
            
        move_mouse_to_safe_position()
        start_calendar_checker()
        if not safe_sleep(1):
            return
            
        move_mouse_to_safe_position()
        pyautogui.press('up')
        if not safe_sleep(1):
            return
        select_window(1)
        
    debug_print("連續任務執行完成")

def download_current_list():
    """只下載當前列表的檔案"""
    global should_stop, stop_event
    
    # 重置停止標誌
    should_stop = False
    stop_event.clear()
    
    debug_print("開始下載當前列表檔案...")
    
    # 啟動按鍵監聽執行緒
    esc_thread = threading.Thread(target=check_esc_key, daemon=True)
    esc_thread.start()
    
    # 將滑鼠移到螢幕中央
    screen_width, screen_height = pyautogui.size()
    pyautogui.moveTo(screen_width // 2, screen_height // 2)
    time.sleep(0.5)
    
    # 執行下載
    select_window(1)

def main():
    try:
        debug_print("=== 自動下載程式 ===")
        debug_print("按下 CTRL+Q 或 CTRL+E 開始連續下載任務")
        debug_print("按下 CTRL+D 下載當前列表檔案")
        debug_print("按下 CTRL+G 檢測檔案列表區域")
        debug_print("按下 ESC 停止下載")
        debug_print("請確保 DOstocksBiz.exe 已開啟且視窗可見")
        
        # 註冊快捷鍵
        keyboard.add_hotkey('ctrl+q', execute_sequence)
        keyboard.add_hotkey('ctrl+e', execute_sequence)
        keyboard.add_hotkey('ctrl+d', download_current_list)
        keyboard.add_hotkey('ctrl+g', start_list_area_checker)
        
        # 初始化排程器
        scheduler = Scheduler(execute_sequence)
        scheduler_thread = scheduler.init_scheduler()
        debug_print("已啟動排程器，將在每天上午 10:00 及 02:24 自動執行下載任務")
        
        # 保持程式運行，但允許 Ctrl+C 中斷
        keyboard.wait('ctrl+c')
            
    except KeyboardInterrupt:
        debug_print("\n程式已結束")
    finally:
        keyboard.unhook_all()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        debug_print("\n程式已結束")
    except Exception as e:
        debug_print(f"程式發生未預期的錯誤: {str(e)}")
