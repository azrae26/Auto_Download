import keyboard
import pywinauto
import pyautogui
import time
from pywinauto.application import Application
import sys
import psutil
import win32gui
import win32con

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
    global selected_window
    target_windows = find_window_handle("stocks")
    if 1 <= index <= len(target_windows):
        selected_window = target_windows[index - 1]
        print(f"\n已選擇視窗: {selected_window[1]}")
        download_files()
    else:
        print("無效的選擇")

def download_files():
    try:
        global selected_window
        
        # 檢查程式是否運行
        if not is_process_running("DOstocksBiz.exe"):
            print("錯誤: DOstocksBiz.exe 未運行，請先開啟程式")
            return

        # 如果還沒有選擇視窗，顯示可用的視窗列表
        if not selected_window:
            target_windows = find_window_handle("stocks")
            
            if not target_windows:
                print("錯誤: 找不到相關視窗")
                return

            print("\n找到以下視窗:")
            for i, (_, title) in enumerate(target_windows, 1):
                print(f"{i}. {title}")
            print("\n請按數字鍵 1-{} 選擇正確的視窗".format(len(target_windows)))
            return

        hwnd, window_title = selected_window

        # 將視窗帶到前景
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
        except Exception as e:
            print(f"警告: 無法將視窗帶到前景: {str(e)}")

        # 連接到程式
        app = Application(backend="uia").connect(handle=hwnd)
        main_window = app.window(handle=hwnd)
        
        # 獲取所有檔案元素
        print("正在掃描檔案列表...")
        files = main_window.descendants(control_type="ListItem")
        
        if not files:
            print("警告: 沒有找到可下載的檔案")
            return

        print(f"找到 {len(files)} 個檔案")
        
        # 設置運行標誌
        running = True
        
        def on_esc():
            nonlocal running
            running = False
            print("程式已暫停")
            keyboard.unhook_all()
            sys.exit()
        
        # 註冊ESC鍵監聽，改用 keyboard.add_hotkey
        keyboard.add_hotkey('esc', on_esc)
        
        # 遍歷並下載所有檔案
        for i, file in enumerate(files, 1):
            if not running:
                print("\n下載已中止")
                keyboard.unhook_all()
                return
                
            try:
                # 獲取檔案位置並雙擊
                rect = file.rectangle()
                center_x = (rect.left + rect.right) // 2
                center_y = (rect.top + rect.bottom) // 2
                
                print(f"正在下載 ({i}/{len(files)}): {file.window_text()}")
                pyautogui.doubleClick(center_x, center_y)
                
                # 縮短等待時間，讓程式更快響應 ESC
                time.sleep(0.5)
            except Exception as e:
                print(f"下載檔案時發生錯誤: {str(e)}")
                continue
            
        print("所有檔案下載完成")
        keyboard.unhook_all()
        
    except Exception as e:
        print(f"發生錯誤: {str(e)}")
        print("\n請確保:")
        print("1. DOstocksBiz.exe 已經開啟")
        print("2. 視窗未最小化")
        print("3. 已選擇正確的視窗")
        keyboard.unhook_all()

def main():
    try:
        print("=== 自動下載程式 ===")
        print("按下 CTRL+Q 或 CTRL+E 開始下載")
        print("按下 ESC 暫停程式")
        print("請確保 DOstocksBiz.exe 已開啟且視窗可見")
        
        # 註冊快捷鍵
        keyboard.add_hotkey('ctrl+q', download_files)
        keyboard.add_hotkey('ctrl+e', download_files)
        keyboard.add_hotkey('esc', lambda: keyboard.unhook_all())
        
        # 註冊數字鍵 1-9 用於選擇視窗
        for i in range(1, 10):
            keyboard.add_hotkey(str(i), lambda x=i: select_window(x))
        
        # 保持程式運行，但允許 Ctrl+C 中斷
        while True:
            time.sleep(0.1)
            
    except KeyboardInterrupt:
        print("\n程式已結束")
        keyboard.unhook_all()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n程式已結束")
    except Exception as e:
        print(f"程式發生未預期的錯誤: {str(e)}")
