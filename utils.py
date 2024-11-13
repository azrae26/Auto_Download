from datetime import datetime
from queue import Queue
import win32gui
import win32con
import time
from pywinauto.application import Application as PywinautoApp
import win32api
import pyautogui

# 全域常數
SLEEP_INTERVAL = 0.1  # 基本等待時間

# 全域變數
debug_queue = Queue()
refresh_checking = False  # 用於控制刷新檢測的開關
last_mouse_pos = None
is_program_moving = False

def debug_print(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f"[{timestamp}] {message}"
    debug_queue.put(formatted_message)
    print(formatted_message)

def find_window_handle(target_title=None):
    """
    找到視窗，使用模糊比對
    target_title: 目標視窗標題（不含版本號）
    """
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # 如果有指定目標標題，使用模糊比對
            if target_title:
                # 移除版本號部分再比對
                window_name = title.split('版本')[0].strip()
                if target_title.lower() in window_name.lower():
                    windows.append((hwnd, title))
            else:
                if title:
                    windows.append((hwnd, title))
    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows 

def ensure_foreground_window(hwnd, window_title=None, sleep_time=0.2):
    """
    確保視窗在前景
    hwnd: 視窗句柄
    window_title: 視窗標題（用於日誌）
    sleep_time: 等待時間（秒）
    """
    try:
        # 檢查視窗是否最小化
        if win32gui.IsIconic(hwnd):
            window_info = f" '{window_title}'" if window_title else ""
            debug_print(f"視窗{window_info}已最小化，正在還原...")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(sleep_time)
        
        # 獲取當前前景視窗
        current_hwnd = win32gui.GetForegroundWindow()
        
        # 如果已經是前景視窗，直接返回
        if current_hwnd == hwnd:
            return True
            
        # 嘗試將視窗帶到前景
        try:
            # 先激活視窗
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            time.sleep(sleep_time)
            
            # 模擬 Alt 鍵按下和釋放，這可以幫助切換視窗焦點
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)  # Alt 按下
            win32gui.SetForegroundWindow(hwnd)
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)  # Alt 釋放
            
            time.sleep(sleep_time)
            return True
            
        except Exception as e:
            debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")
            return False
            
    except Exception as e:
        debug_print(f"確保視窗可見時發生錯誤: {str(e)}")
        return False

def get_list_items(main_window, list_index=None):
    """獲取列表項目"""
    try:
        if list_index is not None:
            return main_window.child_window(
                control_type="List", 
                found_index=list_index
            ).descendants(control_type="ListItem")
        return main_window.descendants(control_type="ListItem")
    except Exception as e:
        debug_print(f"獲取列表項目時發生錯誤: {str(e)}")
        return []

def calculate_center_position(rect):
    """
    計算矩形中心點位置並處理所有可能的錯誤
    返回: (center_x, center_y) 或在出錯時返回 None, None
    """
    try:
        if not rect:
            debug_print("錯誤: 無效的矩形區域")
            return None, None
            
        if not hasattr(rect, 'left') or not hasattr(rect, 'right') or \
           not hasattr(rect, 'top') or not hasattr(rect, 'bottom'):
            debug_print("錯誤: 矩形區域缺少必要屬性")
            return None, None
            
        center_x = (rect.left + rect.right) // 2
        center_y = (rect.top + rect.bottom) // 2
        
        if not isinstance(center_x, (int, float)) or not isinstance(center_y, (int, float)):
            debug_print("錯誤: 計算結果無效")
            return None, None
            
        return center_x, center_y
        
    except Exception as e:
        debug_print(f"計算中心點位置時發生錯誤: {str(e)}")
        return None, None

# 檢查列表是否刷新，但只能檢查出檔案數量變化
def check_list_refresh(current_count, new_count, is_date_switching):
    """
    檢查列表是否發生刷新
    current_count: 當前檔案數量
    new_count: 新的檔案數量
    is_date_switching: 是否正在切換日期
    返回: (是否刷新, 新的檔案數量)
    """
    try:
        if is_date_switching:
            return False, new_count
        
        if current_count != 0 and new_count != current_count:
            debug_print(f"檢測到列表刷新: 檔案數量從 {current_count} 變為 {new_count}")
            return True, new_count
        
        if current_count == 0:
            return False, new_count
        
        return False, current_count
            
    except Exception as e:
        debug_print(f"檢查列表刷新時發生錯誤: {str(e)}")
        return False, current_count

def start_refresh_check(hwnd=None):
    """
    開始檢測列表刷新
    hwnd: 視窗句柄，如果未提供則自動尋找視窗
    """
    global refresh_checking
    CHECK_INTERVAL = 0.3 # 檢查間隔
    
    try:
        if hwnd is None:
            windows = find_window_handle("stocks")
            if not windows:
                debug_print("錯誤: 找不到目標視窗")
                return
            hwnd = windows[0][0]
            debug_print(f"已自動選擇視窗: {windows[0][1]}")
        
        refresh_checking = True
        # 為每個列表維護單獨的檔案集合
        current_files = [{}, {}, {}]  # 使用字典來存儲每個列表的檔案
        last_check_time = datetime.now()
        debug_print("開始檢測列表刷新...")
        
        while refresh_checking:
            try:
                app = PywinautoApp(backend="uia").connect(handle=hwnd)
                main_window = app.window(handle=hwnd)
                
                # 檢查每個列表
                for i in range(3):
                    try:
                        list_control = main_window.child_window(control_type="List", found_index=i)
                        if not list_control.is_visible():
                            continue
                            
                        # 獲取當前列表的檔案
                        files = get_list_items(main_window, i)
                        new_files = {}
                        
                        # 建立新的檔案字典，包含檔案名稱和位置
                        for file in files:
                            try:
                                name = file.window_text()
                                rect = file.rectangle()
                                new_files[name] = (rect.top, rect.bottom)
                            except:
                                continue
                        
                        # 比較檔案變化
                        if current_files[i]:  # 如果有之前的記錄
                            # 檢查檔案位置是否變化
                            for name, pos in new_files.items():
                                if name in current_files[i]:
                                    old_pos = current_files[i][name]
                                    if old_pos != pos:  # 如果位置改變
                                        debug_print(f"檢測到列表 {i+1} 刷新 (檔案位置變化)")
                                        break
                        
                        current_files[i] = new_files
                            
                    except Exception as e:
                        continue
                
            except Exception as e:
                debug_print(f"檢測過程發生錯誤: {str(e)}")
            
            time.sleep(CHECK_INTERVAL)
            
    except Exception as e:
        debug_print(f"檢測列表刷新時發生錯誤: {str(e)}")
    finally:
        refresh_checking = False

def stop_refresh_check():
    """停止檢測列表刷新"""
    global refresh_checking
    refresh_checking = False
    debug_print("停止檢測列表刷新")

def check_mouse_movement():
    """檢查滑鼠是否移動"""
    global last_mouse_pos, is_program_moving
    
    # 如果是程式移動滑鼠，則不檢測
    if is_program_moving:
        return False
        
    current_pos = pyautogui.position()
    
    if last_mouse_pos is None:
        last_mouse_pos = current_pos
        return False
        
    # 檢查是否移動超過 3 像素
    has_moved = (abs(current_pos.x - last_mouse_pos.x) >= 3 or 
                abs(current_pos.y - last_mouse_pos.y) >= 3)
    
    # 只有在非程式移動且確實檢測到移動時才更新位置
    if has_moved and not is_program_moving:
        last_mouse_pos = current_pos
        
    return has_moved

def move_to_safe_position():
    """移動到安全位置"""
    screen_width, screen_height = pyautogui.size()
    pyautogui.moveTo(screen_width // 2, screen_height // 2)
    time.sleep(SLEEP_INTERVAL)  # 使用本地定義的 SLEEP_INTERVAL

def click_at(x, y, is_first_click=False, clicks=1, interval=SLEEP_INTERVAL, sleep_interval=None, hwnd=None, window_title=None, target_element=None):
    """點擊指定位置，若滑鼠移動則暫停1秒"""
    global is_program_moving, last_mouse_pos
    
    try:
        while True:
            if hwnd is None:
                debug_print("警告: 未提供視窗句柄")
                return False

            if window_title is None:
                window_title = win32gui.GetWindowText(hwnd)

            if not ensure_foreground_window(hwnd, window_title):
                debug_print("警告: 無法確保視窗在前景")
                return False

            # 移動滑鼠前檢查
            if check_mouse_movement():
                debug_print("檢測到滑鼠移動，暫停 1 秒")
                time.sleep(1)
                continue
                
            # 設定標記為程式移動
            is_program_moving = True
            
            # 移動滑鼠
            pyautogui.moveTo(x, y)
            time.sleep(interval)
            
            # 更新最後位置為目標位置並重設標記
            last_mouse_pos = pyautogui.position()
            is_program_moving = False
            
            # 檢查滑鼠是否在目標元素上
            if target_element:
                element_rect = target_element.rectangle()
                current_pos = pyautogui.position()
                if not (element_rect.left <= current_pos.x <= element_rect.right and 
                       element_rect.top <= current_pos.y <= element_rect.bottom):
                    debug_print("滑鼠未在目標元素上，重試...")
                    continue
            
            # 點擊前再次檢查
            if check_mouse_movement():
                debug_print("檢測到滑鼠移動，暫停 1 秒")
                time.sleep(1)
                continue
                
            # 執行點擊
            click_completed = True
            for _ in range(clicks):
                pyautogui.click()
                if clicks > 1:
                    time.sleep(interval)
                    # 檢查是否仍在目標元素上
                    if target_element:
                        current_pos = pyautogui.position()
                        element_rect = target_element.rectangle()
                        if not (element_rect.left <= current_pos.x <= element_rect.right and 
                               element_rect.top <= current_pos.y <= element_rect.bottom):
                            debug_print("點擊過程中滑鼠離開目標元素，重試...")
                            click_completed = False
                            break
                        
                    if check_mouse_movement():
                        debug_print("檢測到滑鼠移動，暫停 1 秒")
                        time.sleep(1)
                        click_completed = False
                        break
                        
            if not click_completed:
                continue
            
            if sleep_interval is None:
                sleep_interval = interval * (5 if is_first_click else 1)
            
            time.sleep(sleep_interval)
            return True
            
    except Exception as e:
        is_program_moving = False
        debug_print(f"點擊操作時發生錯誤: {str(e)}")
        return False