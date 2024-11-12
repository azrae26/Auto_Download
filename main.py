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
from datetime import datetime
from calendar_checker import start_calendar_checker
from get_list_area import start_list_area_checker

# 全域變數
should_stop = False
stop_event = threading.Event()
debug_queue = Queue()

pyautogui.FAILSAFE = False

def debug_print(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f"[{timestamp}] {message}"
    debug_queue.put(formatted_message)
    print(formatted_message)

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

def download_files():
    try:
        global selected_window, should_stop
        
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
        list_bottom = main_list_area.bottom
        list_top = main_list_area.top
        
        # 遍歷並下載所有檔案
        i = 0
        while i < len(files):
            if should_stop:
                debug_print("[DEBUG] 檔案迴圈中檢測到停止信號")
                return
                
            try:
                # 每次點擊前確保視窗可見
                if not ensure_window_visible(hwnd, window_title):
                    debug_print("警告: 無法確保視窗可見，重試當前檔案")
                    continue
                
                file = files[i]
                file_name = file.window_text()
                
                # 檢查檔名是否以"_公司"結尾
                if file_name.endswith("_公司"):
                    debug_print(f"跳過檔案 ({i+1}/{len(files)}): {file_name} (檔名以_公司結尾)")
                    i += 1
                    continue
                
                # 獲取檔案位置
                rect = file.rectangle()
                center_x = (rect.left + rect.right) // 2
                center_y = (rect.top + rect.bottom) // 2
                
                # 檢查是否需要翻頁
                if center_y >= (list_bottom - 30):  # 預留一些邊距
                    debug_print("檔案已到達可視範圍底部，需要翻頁")
                    
                    # 確保視窗有焦點
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)  # 等待焦點設置完成
                    except Exception as e:
                        debug_print(f"警告: 無法設置視窗焦點: {str(e)}")
                    
                    # 執行翻頁
                    pyautogui.press('pagedown')
                    time.sleep(0.5)  # 等待翻頁完成
                    
                    # 重新獲取檔案位置
                    rect = file.rectangle()
                    center_x = (rect.left + rect.right) // 2
                    center_y = (rect.top + rect.bottom) // 2
                    
                    # 如果檔案位置沒有改變，可能已經是最後一頁
                    if center_y >= (list_bottom - 30):
                        debug_print("已經是最後一頁")
                
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
                time.sleep(0.05)  # 從0.1改為0.05
                
                # 執行點擊並記錄位置
                pyautogui.doubleClick()
                last_click_x = center_x
                last_click_y = center_y
                
                if not safe_sleep(0.25):  # 從0.5改為0.25
                    return
                
                i += 1
                
            except Exception as e:
                debug_print(f"下載檔案時發生錯誤: {str(e)}")
                i += 1
                continue
            
        debug_print("所有檔案下載完成")
        
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
    global should_stop, stop_event
    
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
    move_mouse_to_safe_position()
    start_calendar_checker()
    if not safe_sleep(1):
        return
        
    # 2. 下載檔案
    debug_print("步驟2: 下載檔案")
    move_mouse_to_safe_position()
    select_window(1)
    if not safe_sleep(1):
        return
        
    # 3. 再次點擊今日
    debug_print("步驟3: 再次點擊今日")
    move_mouse_to_safe_position()
    start_calendar_checker()
    if not safe_sleep(1):
        return
        
    # 4. 如果是週二到週五，執行額外步驟
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

def main():
    try:
        debug_print("=== 自動下載程式 ===")
        debug_print("按下 CTRL+Q 或 CTRL+E 開始連續下載任務")
        debug_print("按下 CTRL+G 檢測檔案列表區域")
        debug_print("按下 ESC 停止下載")
        debug_print("請確保 DOstocksBiz.exe 已開啟且視窗可見")
        
        # 註冊快捷鍵
        keyboard.add_hotkey('ctrl+q', execute_sequence)
        keyboard.add_hotkey('ctrl+e', execute_sequence)
        keyboard.add_hotkey('ctrl+g', start_list_area_checker)
        
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
