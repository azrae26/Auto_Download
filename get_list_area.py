import win32gui
from pywinauto.application import Application as PywinautoApp
from datetime import datetime
import time
from utils import debug_print, find_window_handle, ensure_foreground_window, get_list_items

TARGET_WINDOW = "stocks"  # 直接定義常數

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

def get_list_area():
    try:
        global should_stop
        reset_stop()  # 重置停止標誌
        
        # 尋找目標視窗
        target_windows = find_window_handle(TARGET_WINDOW)
        if not target_windows:
            debug_print("錯誤: 找不到相關視窗")
            return None

        hwnd, window_title = target_windows[0]
        debug_print(f"使用視窗: {window_title}")

        if check_stop():
            debug_print("檢測已停止")
            return None

        if not ensure_foreground_window(hwnd, window_title):
            debug_print("警告: 無法確保視窗在前景")

        if check_stop():
            debug_print("檢測已停止")
            return None

        # 連接到應用程式
        app = PywinautoApp(backend="uia").connect(handle=hwnd)
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
                items = get_list_items(list_area)
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
    """開始檢查列表區域"""
    try:
        debug_print("\n開始尋找所有列表區域:")
        
        # 獲取目標視窗
        target_windows = find_window_handle(TARGET_WINDOW)
        if not target_windows:
            debug_print("找不到目標視窗")
            return None
            
        hwnd = target_windows[0][0]
        app = PywinautoApp(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        
        # 獲取所有列表控件
        all_lists = main_window.descendants(control_type="List")
        
        # 顯示所有找到的列表資訊，包括空白列表
        for i, lst in enumerate(all_lists, 1):
            rect = lst.rectangle()
            debug_print(f"找到列表區域 {i}:")
            debug_print(f"位置: 左={rect.left}, 上={rect.top}, 右={rect.right}, 下={rect.bottom}")
            debug_print(f"寬度: {rect.width()}, 高度: {rect.height()}")
            
            # 獲取列表中的所有項目
            try:
                items = lst.items()
                debug_print("所有項目內容:")
                if items:
                    for j, item in enumerate(items, 1):
                        debug_print(f"  {j}. {item.window_text()}")
                else:
                    debug_print("  [空白列表]")
            except Exception as e:
                debug_print(f"  無法獲取項目: {str(e)}")
            
            debug_print("---")
        
        debug_print(f"\n總共找到 {len(all_lists)} 個列表區域")
        
        if len(all_lists) < 3:
            debug_print("警告: 無法獲取完整的列表區域資訊")
            
        return all_lists
        
    except Exception as e:
        debug_print(f"檢查列表區域時發生錯誤: {str(e)}")
        return None

if __name__ == "__main__":
    start_list_area_checker()