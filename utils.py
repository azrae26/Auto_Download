from datetime import datetime
from queue import Queue
import win32gui
import win32con
import time
from pywinauto.application import Application as PywinautoApp
import win32api
import pyautogui
from config import Config, COLORS  # 添加這行

# 全域常數
SLEEP_INTERVAL = Config.SLEEP_INTERVAL  # 基本等待時間

# 全域變數
debug_queue = Queue() # 用於儲存 debug 訊息的佇列
refresh_checking = False  # 用於控制刷新檢測的開關
last_mouse_pos = None # 用於儲存滑鼠最後位置
is_program_moving = False # 用於控制程式是否移動

def debug_print(message, color='white', bg_color=None, bold=False):
    """
    帶時間戳和顏色的輸出函數
    Args:
        message (str): 要輸出的信息
        color (str): 文字顏色
        bg_color (str, optional): 背景顏色
        bold (bool): 是否粗體
    """
    timestamp = datetime.now().strftime('[%Y-%m-%d %H:%M:%S]')
    color_code = COLORS.get(color, COLORS['reset'])
    bg_code = COLORS.get(bg_color, '') if bg_color else ''
    bold_code = COLORS.get('bold', '') if bold else ''
    print(f"{timestamp} {color_code}{bg_code}{bold_code}{message}{COLORS['reset']}")

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
                # 檢查標題是否包含 "DostocksBiz"
                if "DostocksBiz" in title:
                    debug_print(f"找到視窗: {title}", color='light_green')
                    windows.append((hwnd, title))
            else:
                if title:
                    windows.append((hwnd, title))
    windows = []
    win32gui.EnumWindows(callback, windows)
    
    if not windows:
        debug_print("找不到任何符合的視窗", color='light_red')
        debug_print("請確認 DostocksBiz 程式是否已啟動", color='yellow')
    
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
            debug_print(f"視窗{window_info}已最小化，正在還原...", color='light_cyan')
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
        debug_print(f"確保視窗可見時發生錯誤: {str(e)}", color='light_red')
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
                    debug_print(f"從 {list_type} 列表獲取到 {len(items)} 個項目", color='light_blue', bold=True)
                    all_items.extend(items)
                except Exception as e:
                    debug_print(f"獲取 {list_type} 列表項目時發生錯誤: {str(e)}", color='light_red')
                    continue
            return all_items
            
    except Exception as e:
        debug_print(f"獲取列表項目時發生錯誤: {str(e)}", color='light_red')
        return []

def calculate_center_position(rect):
    """
    計算矩形中心點位置並處理所有可能的錯誤
    返回: (center_x, center_y) 或在出錯時返回 None, None
    """
    try:
        if not rect:
            debug_print("錯誤: 無效的矩形區域", color='light_red')
            return None, None
            
        if not hasattr(rect, 'left') or not hasattr(rect, 'right') or \
           not hasattr(rect, 'top') or not hasattr(rect, 'bottom'):
            debug_print("錯誤: 矩形區域缺少必要屬性", color='light_red')
            return None, None
            
        center_x = (rect.left + rect.right) // 2
        center_y = (rect.top + rect.bottom) // 2
        
        if not isinstance(center_x, (int, float)) or not isinstance(center_y, (int, float)):
            debug_print("錯誤: 計算結果無效", color='light_red')
            return None, None
            
        return center_x, center_y
        
    except Exception as e:
        debug_print(f"計算中心點位置時發生錯誤: {str(e)}", color='light_red')
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
        debug_print("開始檢測列表刷新...", color='light_cyan')
        
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
                            debug_print(f"檢測到列表 {i+1} 刷新", color='light_green')
                            
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
        debug_print(f"檢查元素時發生錯誤: {str(e)}", color='light_red')
        return False

def click_at(x, y, is_first_click=False, clicks=1, interval=SLEEP_INTERVAL, sleep_interval=None, hwnd=None, window_title=None, expected_text=None):
    """使用 win32api 進行點擊，並檢查元素文字"""
    global is_program_moving, last_mouse_pos
    
    try: 
        max_retries = 2  # 最大重試次數
        retry_count = 0
        
        # 設定標記為程式移動 - 移到這裡，整個操作開始前就設置
        is_program_moving = True
        
        try:
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
                    
                win32api.SetCursorPos((x, y))
                time.sleep(interval * 0.1)
                
                # 更新位置
                current_pos = pyautogui.position()
                last_mouse_pos = current_pos
                
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
                    
                    success = True
                    
                    if clicks > 1:
                        time.sleep(interval * 3)
                
                if success:
                    # 點擊後等待時間，第一次點擊等待時間
                    time.sleep(sleep_interval or (interval * (5 if is_first_click else 0.5)))
                    return True
                else:
                    retry_count += 1
                    continue
                    
        finally:
            is_program_moving = False  # 確保標記被重置
            
        debug_print(f"已達最大重試次數 {max_retries}，放棄此次點擊", color='light_red')
        return False
            
    except Exception as e:
        is_program_moving = False  # 確保發生錯誤時重設標記
        debug_print(f"點擊操作時發生錯誤: {str(e)}", color='light_red')
        return False

def scroll_to_file(file, list_area, hwnd):
    """滾動直到檔案進入可視範圍"""
    try:
        debug_print("開始滾動檢查...", color='light_cyan')
        
        # 連接到視窗
        app = PywinautoApp(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        
        # 獲取視窗標題
        window_title = win32gui.GetWindowText(hwnd)
        
        # 獲取目標檔案名稱
        target_name = file.window_text()
        debug_print(f"目標檔案: {target_name}", color='light_blue', bold=True)

        # 獲取三個列表的檔案
        morning_files = get_list_items_by_id(main_window, 'morning')
        research_files = get_list_items_by_id(main_window, 'research')
        industry_files = get_list_items_by_id(main_window, 'industry')
        
        # 找出目標檔案在哪個列表及其序號
        target_list = None
        target_index = None
        target_list_id = None
        
        for list_type, (list_id, files) in [
            ('晨會報告', ('listBoxMorningReports', morning_files)), 
            ('研究報告', ('listBoxResearchReports', research_files)), 
            ('產業報告', ('listBoxIndustryReports', industry_files))
        ]:
            try:
                index = [f.window_text() for f in files].index(target_name)
                target_list = list_type
                target_index = index
                target_list_id = list_id
                debug_print(f"目標檔案在 [{list_type}] 列表，序號: {index}", color='light_blue', bold=True)
                break
            except ValueError:
                continue
                
        if target_index is None:
            debug_print("找不到目標檔案", color='light_red')
            return False
            
        # 找出游標選中的檔案在哪個列表及其序號
        current_list = None
        current_index = None
        
        for list_type, files in [('晨會報告', morning_files), 
                               ('研究報告', research_files), 
                               ('產業報告', industry_files)]:
            for i, f in enumerate(files):
                if f.is_selected():
                    current_list = list_type
                    current_index = i
                    debug_print(f"游標選中的檔案在 [{list_type}] 列表，序號: {i}", color='light_blue', bold=True)
                    break
            if current_index is not None:
                break
                
        if current_index is None:
            debug_print("找不到游標選中的檔案", color='light_red')
            # 將中文列表名稱轉換為英文標識符
            list_type = None
            if target_list == '晨會報告':
                list_type = 'morning'
            elif target_list == '研究報告':
                list_type = 'research'
            elif target_list == '產業報告':
                list_type = 'industry'
                
            if list_type:
                debug_print(f"嘗試點擊 [{target_list}] 列表", color='yellow')
                # 點擊切換到目標檔案所在的列表
                if not switch_to_list(hwnd, list_type, next_list=False):
                    debug_print("切換到目標列表失敗", color='light_red')
                    return False
                    
                # 重新獲取目標列表區域
                list_area = main_window.child_window(auto_id=target_list_id)
                debug_print("已更新列表區域", color='light_green')
                
                # 重新檢查檔案可見性
                if is_file_visible(file, list_area):
                    debug_print("切換列表後檔案已可見", color='light_green')
                    return True
                    
                # 設定當前列表和索引
                current_list = target_list
                current_index = target_index  # 設定為目標檔案的索引
                debug_print(f"更新當前索引為: {current_index}", color='light_magenta')
            else:
                debug_print("無法識別目標列表類型", color='light_red')
                return False
                
        # 檢查是否需要切換列表
        if target_list != current_list:
            debug_print(f"需要從 [{current_list}] 切換到 [{target_list}]", color='light_blue', bold=True)
            # 切換列表時傳入視窗標題
            if not switch_to_list(hwnd):
                return False
            time.sleep(0.2)  # 等待列表切換完成
            
            # 重新獲取目標列表區域
            list_area = main_window.child_window(auto_id=target_list_id)
            debug_print("已更新列表區域", color='light_green')
            
            # 重新檢查檔案可見性
            if is_file_visible(file, list_area):
                debug_print("切換列表後檔案已可見", color='light_green')
                return True
        
        # 設定最大嘗試次數
        max_attempts = 10
        attempts = 0
        
        # 在開始翻頁前，確保目標列表被選中
        list_type = None
        if target_list == '晨會報告':
            list_type = 'morning'
        elif target_list == '研究報告':
            list_type = 'research'
        elif target_list == '產業報告':
            list_type = 'industry'
            
        if list_type:
            debug_print(f"翻頁前確保 [{target_list}] 列表被選中", color='light_blue', bold=True)
            if not switch_to_list(hwnd, list_type, next_list=False):
                debug_print("切換到目標列表失敗", color='light_red')
                return False
            time.sleep(0.1)  # 等待列表切換完成
        
        # 執行翻頁直到找到檔案
        while attempts < max_attempts:
            if is_file_visible(file, list_area):
                debug_print("目標檔案已可見", color='light_green')
                return True
                
            # 確保視窗在前景
            if not ensure_foreground_window(hwnd, window_title):
                debug_print("視窗不在前景，重新嘗試滾動", color='light_red')
                attempts += 1
                continue
                
            # 決定滾動方向
            if target_index < current_index:
                debug_print("目標檔案在上方，向上翻頁", color='yellow')
                pyautogui.press('pageup')
            else:
                debug_print("目標檔案在下方，向下翻頁", color='yellow')
                pyautogui.press('pagedown')
                
            time.sleep(0.2)
            
            # 檢查檔案是否可見
            if is_file_visible(file, list_area):
                debug_print("翻頁後檔案已可見", color='light_green')
                return True
                
            attempts += 1
            
        debug_print(f"已達最大嘗試次數 {max_attempts}，無法找到目標檔案", color='light_red')
        return False
                
    except Exception as e:
        debug_print(f"滾動到檔案位置時發生錯誤: {str(e)}", color='light_red')
        return False

def is_file_visible(file, list_area):
    """檢查檔案是否在可視範圍內"""
    try:
        # 獲取檔案和列表的矩形尺寸及座標
        file_rect = file.rectangle()
        list_rect = list_area.rectangle()
        
        # 檢查檔案頂部是否在可視範圍內
        top_visible = file_rect.top >= list_rect.top
        
        # 檢查檔案底部是否在可視範圍內
        bottom_visible = file_rect.bottom <= list_rect.bottom
        
        # 檢查檔案是否可見
        is_visible = top_visible and bottom_visible

        file_name = file.window_text()
        
        if not is_visible:
            debug_print(f"{file_name} 檔案不在可視範圍內 (top={file_rect.top}, bottom={file_rect.bottom})", color='light_magenta')
            debug_print(f"列表範圍: top={list_rect.top}, bottom={list_rect.bottom}", color='light_magenta')
            
        return is_visible
        
    except Exception as e:
        debug_print(f"檢查檔案可見性時發生錯誤: {str(e)}", color='light_red')
        return False

def switch_to_list(hwnd, list_type=None, next_list=True):
    """
    點擊切換到指定列表或下一個列表
    hwnd: 視窗句柄
    list_type: 'morning'|'research'|'industry' 指定要切換到哪個列表
    next_list: True=切換到下一個列表, False=切換到指定列表
    """
    global is_program_moving  # 添加這行
    
    try:
        # 連接到視窗
        app = PywinautoApp(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        
        list_types = {
            'morning': ('listBoxMorningReports', '晨會報告'),
            'research': ('listBoxResearchReports', '研究報告'),
            'industry': ('listBoxIndustryReports', '產業報告')
        }

        # 標記為程式移動
        is_program_moving = True
        
        try:
            if next_list:
                # 切換到下一個列表的邏輯
                lists = main_window.descendants(control_type="List")
                if len(lists) >= 2:
                    next_list = lists[1]  # 從左側列表切換到中間列表
                    rect = next_list.rectangle()
                    debug_print("切換到下一個列表", color='yellow')
                else:
                    debug_print("警告: 找不到足夠的列表區域", color='light_red')
                    return False
            else:
                # 切換到指定列表的邏輯
                if list_type not in list_types:
                    debug_print(f"錯誤: 無效的列表類型 {list_type}", color='light_red')
                    return False
                    
                list_id, list_name = list_types[list_type]
                target_list = main_window.child_window(auto_id=list_id)
                rect = target_list.rectangle()
                debug_print(f"切換到{list_name}列表", color='yellow')

            # 計算點擊位置（列表頂部往下 10px 的位置）
            center_x = (rect.left + rect.right) // 2
            click_y = rect.top + 10
            
            # 移動到位置並點擊
            click_at(
                center_x, 
                click_y, 
                clicks=1, 
                sleep_interval=SLEEP_INTERVAL * 0.2,
                hwnd=hwnd, 
                window_title=win32gui.GetWindowText(hwnd)
            )
            
            debug_print(f"已完成列表切換", color='light_green')
            return True
            
        finally:
            is_program_moving = False  # 確保標記被重置
            
    except Exception as e:
        is_program_moving = False  # 確保發生錯誤時重設標記
        debug_print(f"切換列表時發生錯誤: {str(e)}", color='light_red')
        return False