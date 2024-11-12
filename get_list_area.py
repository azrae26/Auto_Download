import win32gui
from pywinauto.application import Application
from datetime import datetime
import time

def debug_print(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f"[{timestamp}] {message}"
    print(formatted_message)

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

def get_list_area():
    try:
        # 尋找目標視窗
        target_windows = find_window_handle("stocks")
        if not target_windows:
            debug_print("錯誤: 找不到相關視窗")
            return None

        hwnd, window_title = target_windows[0]
        debug_print(f"使用視窗: {window_title}")

        # 將視窗帶到前景
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.2)
        except Exception as e:
            debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")

        # 連接到應用程式
        app = Application(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)

        # 找到所有列表區域
        debug_print("\n開始尋找所有列表區域:")
        list_areas = []
        
        # 方法1: 通過控件類型尋找
        lists = main_window.descendants(control_type="List")
        for list_area in lists:
            try:
                rect = list_area.rectangle()
                # 檢查列表是否有內容
                items = list_area.children(control_type="ListItem")
                if items:
                    list_areas.append(rect)
                    debug_print(f"找到列表區域 {len(list_areas)}:")
                    debug_print(f"位置: 左={rect.left}, 上={rect.top}, 右={rect.right}, 下={rect.bottom}")
                    debug_print(f"寬度: {rect.right - rect.left}, 高度: {rect.bottom - rect.top}")
                    
                    # 顯示前三個項目的內容（如果有的話）
                    debug_print("前三個項目內容:")
                    for i, item in enumerate(items[:3]):
                        try:
                            text = item.window_text()
                            debug_print(f"  {i+1}. {text}")
                        except:
                            pass
                    debug_print("---")
            except:
                continue

        if list_areas:
            debug_print(f"\n總共找到 {len(list_areas)} 個列表區域")
            return list_areas
        else:
            debug_print("警告: 沒有找到任何列表區域")
            return None

    except Exception as e:
        debug_print(f"檢測列表區域時發生錯誤: {str(e)}")
        return None

def start_list_area_checker():
    debug_print("開始檢測檔案列表區域...")
    return get_list_area() 