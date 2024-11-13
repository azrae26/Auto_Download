import win32gui
from pywinauto.application import Application
from datetime import datetime
import time
from utils import debug_print

# 全域變數用於控制停止
should_stop = False

def check_stop():
    """檢查是否應該停止"""
    global should_stop
    return should_stop

def set_stop():
    """設置停止標誌"""
    global should_stop
    should_stop = True

def reset_stop():
    """重置停止標誌"""
    global should_stop
    should_stop = False

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
        global should_stop
        reset_stop()  # 重置停止標誌
        
        # 尋找目標視窗
        target_windows = find_window_handle("stocks")
        if not target_windows:
            debug_print("錯誤: 找不到相關視窗")
            return None

        hwnd, window_title = target_windows[0]
        debug_print(f"使用視窗: {window_title}")

        if check_stop():
            debug_print("檢測已停止")
            return None

        # 將視窗帶到前景
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.2)
        except Exception as e:
            debug_print(f"警告: 無法將視窗帶到前景: {str(e)}")

        if check_stop():
            debug_print("檢測已停止")
            return None

        # 連接到應用程式
        app = Application(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)

        # 找到所有列表區域
        debug_print("\n開始尋找所有列表區域:")
        list_areas = []
        
        # 方法1: 通過控件類型尋找
        lists = main_window.descendants(control_type="List")
        for list_area in lists:
            if check_stop():
                debug_print("檢測已停止")
                return None
                
            try:
                rect = list_area.rectangle()
                # 檢查列表是否有內容
                items = list_area.children(control_type="ListItem")
                if items:
                    list_areas.append(rect)
                    debug_print(f"找到列表區域 {len(list_areas)}:")
                    debug_print(f"位置: 左={rect.left}, 上={rect.top}, 右={rect.right}, 下={rect.bottom}")
                    debug_print(f"寬度: {rect.right - rect.left}, 高度: {rect.bottom - rect.top}")
                    
                    # 顯示所有項目的內容，空的顯示"無"
                    debug_print("所有項目內容:")
                    for i, item in enumerate(items):
                        if check_stop():
                            debug_print("檢測已停止")
                            return None
                        try:
                            text = item.window_text()
                            if text:
                                debug_print(f"  {i+1}. {text}")
                            else:
                                debug_print(f"  {i+1}. 無")
                        except Exception as e:
                            debug_print(f"  {i+1}. 無 (錯誤: {str(e)})")
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

if __name__ == "__main__":
    start_list_area_checker()