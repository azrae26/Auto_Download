from datetime import datetime
from queue import Queue
import win32gui
import win32con
import time
from pywinauto.application import Application as PywinautoApp
import win32api
import pyautogui
from colorama import init, Fore, Style

# 初始化 colorama
init()

# 顏色映射
COLORS = {
    'red': Fore.RED,                      # 錯誤訊息，紅色
    'orange': Fore.LIGHTRED_EX,           # 橘色
    'green': Fore.GREEN,                  # 成功訊息，綠色
    'yellow': Fore.YELLOW,                # 操作訊息，黃色
    'blue': Fore.BLUE,                    # 一般資訊，藍色
    'magenta': Fore.MAGENTA,              # 檢測訊息，洋紅色
    'cyan': Fore.CYAN,                    # 標題，青色
    'light_red': Fore.LIGHTRED_EX,        # 亮紅色
    'light_green': Fore.LIGHTGREEN_EX,    # 亮綠色
    'light_yellow': Fore.LIGHTYELLOW_EX,  # 亮黃色
    'light_blue': Fore.LIGHTBLUE_EX,      # 亮藍色
    'light_magenta': Fore.LIGHTMAGENTA_EX,# 亮洋紅色
    'light_cyan': Fore.LIGHTCYAN_EX,      # 亮青色
    'light_white': Fore.LIGHTWHITE_EX,    # 亮白色
    'dark_grey': Fore.LIGHTBLACK_EX,      # 深灰色
    'reset': Style.RESET_ALL,             # 重置顏色
}

# 全域常數
SLEEP_INTERVAL = 0.1  # 基本等待時間

# 全域變數
debug_queue = Queue()
refresh_checking = False  # 用於控制刷新檢測的開關
last_mouse_pos = None
is_program_moving = False

def debug_print(message, color='white'):
    """帶時間戳和顏色的輸出函數"""
    timestamp = datetime.now().strftime('[%Y-%m-%d %H:%M:%S]')
    color_code = COLORS.get(color, Fore.WHITE)
    print(f"{timestamp} {color_code}{message}{Style.RESET_ALL}")

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
            debug_print(f"視窗{window_info}已最小化，正在還原...", color='yellow')
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
            debug_print(f"警告: 無法將視窗帶到前景: {str(e)}", color='light_red')
            return False
            
    except Exception as e:
        debug_print(f"確保視窗可見時發生錯誤: {str(e)}", color='red')
        return False

def get_list_items_by_id(main_window, list_type=None):
    """
    根據自動化 ID 獲取列表項目
    list_type: 'morning' | 'research' | 'industry' | None
    如果 list_type 為 None，則返回所有列表的項目
    """
    try:
        list_ids = {
            'morning': 'listBoxMorningReports',
            'research': 'listBoxResearchReports',
            'industry': 'listBoxIndustryReports'
        }
        
        if list_type:
            # 獲取特定列表的項目
            list_id = list_ids.get(list_type)
            if not list_id:
                debug_print(f"錯誤: 無效的列表類型 {list_type}", color='light_red')
                return []
                
            list_control = main_window.child_window(auto_id=list_id)
            return list_control.descendants(control_type="ListItem")
        else:
            # 獲取所有列表的項目
            all_items = []
            for list_type, list_id in list_ids.items():
                try:
                    list_control = main_window.child_window(auto_id=list_id)
                    items = list_control.descendants(control_type="ListItem")
                    debug_print(f"從 {list_type} 列表獲取到 {len(items)} 個項目", color='blue')
                    all_items.extend(items)
                except Exception as e:
                    debug_print(f"獲取 {list_type} 列表項目時發生錯誤: {str(e)}", color='light_red')
                    continue
            return all_items
            
    except Exception as e:
        debug_print(f"獲取列表項目時發生錯誤: {str(e)}", color='red')
        return []

def calculate_center_position(rect):
    """
    計算矩形中心點位置並處理所有可能的錯誤
    返回: (center_x, center_y) 或在出錯時返回 None, None
    """
    try:
        if not rect:
            debug_print("錯誤: 無效的矩形區域", color='red')
            return None, None
            
        if not hasattr(rect, 'left') or not hasattr(rect, 'right') or \
           not hasattr(rect, 'top') or not hasattr(rect, 'bottom'):
            debug_print("錯誤: 矩形區域缺少必要屬性", color='red')
            return None, None
            
        center_x = (rect.left + rect.right) // 2
        center_y = (rect.top + rect.bottom) // 2
        
        if not isinstance(center_x, (int, float)) or not isinstance(center_y, (int, float)):
            debug_print("錯誤: 計算結果無效", color='red')
            return None, None
            
        return center_x, center_y
        
    except Exception as e:
        debug_print(f"計算中心點位置時發生錯誤: {str(e)}", color='red')
        return None, None

def start_refresh_check(hwnd=None):
    """開始檢測列表刷新"""
    global refresh_checking
    CHECK_INTERVAL = 0.3
    
    try:
        if hwnd is None:
            windows = find_window_handle("stocks")
            if not windows:
                debug_print("錯誤: 找不到目標視窗", color='light_red')
                return
            hwnd = windows[0][0]
            
        refresh_checking = True
        current_files = [{}, {}, {}]  # 每個列表的檔案狀態
        debug_print("開始檢測列表刷新...", color='yellow')
        
        while refresh_checking:
            try:
                app = PywinautoApp(backend="uia").connect(handle=hwnd)
                main_window = app.window(handle=hwnd)
                
                for i in range(3):
                    try:
                        list_control = main_window.child_window(control_type="List", found_index=i)
                        if not list_control.is_visible():
                            continue
                            
                        new_files = {
                            file.window_text(): file.rectangle()
                            for file in get_list_items_by_id(main_window, i)
                        }
                        
                        if current_files[i] and new_files != current_files[i]:
                            debug_print(f"檢測到列表 {i+1} 刷新", color='green')
                            
                        current_files[i] = new_files
                            
                    except Exception:
                        continue
                        
            except Exception as e:
                debug_print(f"檢測過程發生錯誤: {str(e)}", color='light_red')
                
            time.sleep(CHECK_INTERVAL)
            
    except Exception as e:
        debug_print(f"檢測列表刷新時發生錯誤: {str(e)}", color='light_red')
    finally:
        refresh_checking = False

def stop_refresh_check():
    """停止檢測列表刷新"""
    global refresh_checking
    refresh_checking = False
    debug_print("停止檢測列表刷新", color='yellow')

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


def check_target_element(hwnd, x, y, expected_text=None):
    """檢查指定位置的元素是否符合預期"""
    try:
        # 如果沒有預期文字，則返回 True
        if not expected_text:
            return True
            
        # 連接到視窗
        app = PywinautoApp(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        
        # 搜索範圍（像素）
        search_range = 5
        
        # 在周圍區域搜索元素
        for offset_x in range(-search_range, search_range + 1): # 負數
            for offset_y in range(-search_range, search_range + 1): # 負數
                try:
                    element = main_window.from_point(x + offset_x, y + offset_y) # 正負數
                    if element:
                        element_text = element.window_text() # 文字
                        if element_text == expected_text: # 文字比對
                            return True
                except:
                    continue
                    
        debug_print(f"點擊位置不是目標元素: {expected_text}", color='light_red')
        return False
 
    except Exception as e:
        debug_print(f"檢查元素時發生錯誤: {str(e)}", color='red')
        return False

def click_at(x, y, is_first_click=False, clicks=1, interval=SLEEP_INTERVAL, sleep_interval=None, hwnd=None, window_title=None, expected_text=None):
    """使用 win32api 進行點擊，並檢查元素文字"""
    global is_program_moving, last_mouse_pos
    
    try: 
        max_retries = 3  # 最大重試次數
        retry_count = 0
        
        while retry_count < max_retries:
            # 確保視窗在前景
            if not ensure_foreground_window(hwnd, window_title):
                debug_print("視窗不在前景，重新嘗試點擊", color='light_red')
                retry_count += 1
                continue
            
            # 移動滑鼠前檢查
            if check_mouse_movement():
                debug_print("檢測到滑鼠移動，暫停 1 秒後重試", color='light_red')
                time.sleep(1)
                retry_count += 1
                continue
                
            # 設定標記為程式移動並執行移動
            is_program_moving = True
            win32api.SetCursorPos((x, y))
            time.sleep(interval * 0.1)
            
            # 更新位置並重設標記
            current_pos = pyautogui.position()
            last_mouse_pos = current_pos
            is_program_moving = False
            
            # 計算絕對座標
            normalized_x = int(x * 65535 / win32api.GetSystemMetrics(0))
            normalized_y = int(y * 65535 / win32api.GetSystemMetrics(1))
            
            # 執行點擊
            for _ in range(clicks):
                win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE | win32con.MOUSEEVENTF_MOVE, 
                                   normalized_x, normalized_y, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                
                # 在按下和彈起之間檢查目標元素
                if expected_text and not check_target_element(hwnd, x, y, expected_text):
                    debug_print(f"重新嘗試點擊", color='light_red')
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                    success = False
                    break
                
                # 確保視窗在前景
                if not ensure_foreground_window(hwnd, window_title):
                    debug_print("視窗不在前景，重新嘗試點擊", color='light_red')
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                    success = False
                    break
                
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                
                success = True # 點擊成功

                if clicks > 1:
                    time.sleep(interval * 3)
            
            if success:
                # 點擊後等待時間，第一次點擊等待時間
                time.sleep(sleep_interval or (interval * (5 if is_first_click else 2)))
                return True
            else:
                retry_count += 1
                continue
            
        debug_print(f"已達最大重試次數 {max_retries}，放棄此次點擊", color='light_red')
        return False
            
    except Exception as e:
        is_program_moving = False  # 確保發生錯誤時重設標記
        debug_print(f"點擊操作時發生錯誤: {str(e)}", color='light_red')
        return False

def scroll_to_file(file, list_area, hwnd):
    """滾動直到檔案進入可視範圍"""
    try:
        debug_print("開始滾動檢查...", color='magenta')
        
        # 連接到視窗
        app = PywinautoApp(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        debug_print("已連接視窗", color='magenta')
        
        # 獲取列表區域的矩形
        list_rect = list_area.rectangle()
        debug_print("已獲取列表區域", color='magenta')
        
        # 獲取當前各個列表的所有檔案
        debug_print("開始獲取檔案列表...", color='magenta')
        current_files_morning = get_list_items_by_id(main_window, list_type='morning')
        current_files_research = get_list_items_by_id(main_window, list_type='research')
        current_files_industry = get_list_items_by_id(main_window, list_type='industry')
        
        debug_print("完成獲取檔案列表", color='magenta')
        
        # 獲取目標檔案名稱
        target_name = file.window_text()
        debug_print(f"目標檔案: {target_name}", color='magenta')

        # 判斷目標檔案在列表的序列數
        target_index_morning = current_files_morning.index(file)
        target_index_research = current_files_research.index(file)
        target_index_industry = current_files_industry.index(file)
        debug_print(f"目標檔案在列表的序列數: {target_index_morning}", color='magenta')
        debug_print(f"目標檔案在列表的序列數: {target_index_research}", color='magenta')
        debug_print(f"目標檔案在列表的序列數: {target_index_industry}", color='magenta')

                    
        # 設定最大嘗試次數
        max_attempts = 10
        attempts = 0
        
        # 比較兩者，判斷滾動方向並執行翻頁，直到找到檔案或達到最大嘗試次數
        while attempts < max_attempts:
            # 獲取游標選中檔案的序列數 - 新的實現方式
            try:
                app = PywinautoApp(backend="uia").connect(handle=hwnd)
                main_window = app.window(handle=hwnd)
                
                # 直接獲取所有 ListItem 控件
                all_items = main_window.descendants(control_type="ListItem")
                
                # 找出被選中的項目
                current_index = None
                for i, item in enumerate(all_items):
                    if item.is_selected():
                        current_index = i
                        debug_print(f"當前選中的檔案: {item.window_text()}", color='magenta')
                        debug_print(f"選中檔案的序列數: {current_index}", color='magenta')
                        break
                
                if current_index is None:
                    debug_print("找不到選中的檔案，使用可見檔案判斷", color='yellow')
                    for i, item in enumerate(all_items):
                        if is_file_visible(item, list_area):
                            current_index = i
                            debug_print(f"使用可見檔案的序列數: {current_index}", color='magenta')
                            break
            except Exception as e:
                debug_print(f"獲取選中檔案時發生錯誤: {str(e)}", color='light_red')
                return False

            # 確保視窗在前景
            if not ensure_foreground_window(hwnd, win32gui.GetWindowText(hwnd)):
                debug_print("警告: 無法確保視窗前景", color='light_red')
                return False


            # 檢查檔案是否已經可見
            if is_file_visible(file, list_area):
                debug_print("目標檔案已可見", color='green')
                return True
                
            # 決定滾動方向
            if target_index < current_index:
                debug_print("目標檔案在可見區域上方，向上翻頁", color='orange')
                pyautogui.press('pageup')
            else:
                debug_print("目標檔案在可見區域下方，向下翻頁", color='orange')
                pyautogui.press('pagedown')

            # 等待0.2秒，確保翻頁完成
            time.sleep(0.2)
            
            # 更新當前索引
            for i, f in enumerate(current_files):
                if is_file_visible(f, list_area):
                    current_index = i
                    break
            
            attempts += 1
            
        debug_print(f"已達最大嘗試次數 {max_attempts}，無法找到目標檔案", color='light_red')
        return False
                
    except Exception as e:
        debug_print(f"滾動到檔案位置時發生錯誤: {str(e)}", color='light_red')
        return False

def is_file_visible(file, list_area):
    """檢查檔案是否在可視範圍內"""
    try:
        file_rect = file.rectangle()  # 獲取檔案的矩形
        list_rect = list_area.rectangle()  # 獲取列表區域的矩形
        
        # 檢查是否為列表最後一個檔案
        current_files = list_area.descendants(control_type="ListItem")
        is_last_file = file == current_files[-1] if current_files else False
        
        # 檢查檔案頂部是否在可視範圍內
        top_visible = file_rect.top >= list_rect.top
        
        # 檢查檔案底部是否在可視範圍內（最後一個檔案不加5像素）
        bottom_offset = 0 if is_last_file else 5
        bottom_visible = file_rect.bottom + bottom_offset <= list_rect.bottom
        
        return top_visible and bottom_visible
        
    except Exception as e:
        debug_print(f"檢查檔案可見性時發生錯誤: {str(e)}", color='light_red')
        return False