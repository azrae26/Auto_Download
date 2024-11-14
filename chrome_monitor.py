from datetime import datetime
import time
import win32gui
import win32con
from pywinauto.application import Application
from utils import debug_print
import threading

class ChromeMonitor:
    def __init__(self):
        self.chrome_windows = {}  # 儲存 Chrome 視窗資訊
        self.is_monitoring = False
        self.monitor_thread = None
        
    def find_chrome_windows(self):
        """找到所有的 Chrome 視窗"""
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if " - Google Chrome" in title:
                    windows.append((hwnd, title))
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows
    
    def count_tabs(self, hwnd):
        """計算指定 Chrome 視窗中的分頁數"""
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            chrome_window = app.window(handle=hwnd)
            tab_list = chrome_window.child_window(control_type="TabList")
            
            if tab_list:
                tabs = tab_list.children(control_type="TabItem")
                return len(tabs)
            return 0
            
        except Exception as e:
            debug_print(f"計算分頁數時發生錯誤: {str(e)}")
            return 0
    
    def get_tab_titles(self, hwnd):
        """獲取指定 Chrome 視窗中所有分頁的標題"""
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            chrome_window = app.window(handle=hwnd)
            tab_list = chrome_window.child_window(control_type="TabList")
            
            if tab_list:
                tabs = tab_list.children(control_type="TabItem")
                return [tab.window_text() for tab in tabs]
            return []
            
        except Exception as e:
            debug_print(f"獲取分頁標題時發生錯誤: {str(e)}")
            return []
    
    def check_window_changes(self, hwnd, current_info):
        """檢查視窗的變化並報告"""
        try:
            # 獲取新的資訊
            new_tab_count = self.count_tabs(hwnd)
            new_tab_titles = self.get_tab_titles(hwnd)
            new_pdf_count = sum(1 for title in new_tab_titles if '.pdf' in title.lower())
            new_ppt_count = sum(1 for title in new_tab_titles if any(ext in title.lower() for ext in ['.ppt', '.pptx']))
            
            # 檢查變化
            if new_tab_count != current_info['tab_count']:
                debug_print(f"\n分頁數量變化: {current_info['tab_count']} -> {new_tab_count}")
            
            if new_pdf_count != current_info['pdf_count']:
                debug_print(f"PDF 檔案數量變化: {current_info['pdf_count']} -> {new_pdf_count}")
            
            if new_ppt_count != current_info['ppt_count']:
                debug_print(f"PPT 檔案數量變化: {current_info['ppt_count']} -> {new_ppt_count}")
            
            # 更新資訊
            return {
                'title': current_info['title'],
                'tab_count': new_tab_count,
                'pdf_count': new_pdf_count,
                'ppt_count': new_ppt_count,
                'tab_titles': new_tab_titles,
                'last_update': datetime.now()
            }
            
        except Exception as e:
            debug_print(f"檢查視窗變化時發生錯誤: {str(e)}")
            return current_info

    def monitor_loop(self):
        """持續監控 Chrome 視窗的變化"""
        debug_print("\n開始監控 Chrome 視窗...")
        
        while self.is_monitoring:
            try:
                chrome_windows = self.find_chrome_windows()
                
                for hwnd, title in chrome_windows:
                    # 如果是新視窗，初始化資訊
                    if hwnd not in self.chrome_windows:
                        tab_count = self.count_tabs(hwnd)
                        tab_titles = self.get_tab_titles(hwnd)
                        pdf_count = sum(1 for t in tab_titles if '.pdf' in t.lower())
                        ppt_count = sum(1 for t in tab_titles if any(ext in t.lower() for ext in ['.ppt', '.pptx']))
                        
                        self.chrome_windows[hwnd] = {
                            'title': title,
                            'tab_count': tab_count,
                            'pdf_count': pdf_count,
                            'ppt_count': ppt_count,
                            'tab_titles': tab_titles,
                            'last_update': datetime.now()
                        }
                        debug_print(f"\n新增視窗監控: {title}")
                        debug_print(f"初始分頁數: {tab_count}")
                        debug_print(f"PDF 檔案數: {pdf_count}")
                        debug_print(f"PPT 檔案數: {ppt_count}")
                    else:
                        # 檢查現有視窗的變化
                        self.chrome_windows[hwnd] = self.check_window_changes(
                            hwnd, 
                            self.chrome_windows[hwnd]
                        )
                
                # 移除已關閉的視窗
                closed_windows = set(self.chrome_windows.keys()) - set(hwnd for hwnd, _ in chrome_windows)
                for hwnd in closed_windows:
                    window_info = self.chrome_windows.pop(hwnd)
                    debug_print(f"\n視窗已關閉: {window_info['title']}")
                
                time.sleep(1)  # 每秒檢查一次
                
            except Exception as e:
                debug_print(f"監控過程發生錯誤: {str(e)}")
                time.sleep(1)
    
    def start_monitoring(self):
        """開始監控"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
            self.monitor_thread.start()
            debug_print("Chrome 監控已啟動")
    
    def stop_monitoring(self):
        """停止監控"""
        self.is_monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=1.0)
            debug_print("Chrome 監控已停止")

def start_chrome_monitor(existing_monitor=None):
    """啟動或切換 Chrome 監控"""
    if existing_monitor and existing_monitor.is_monitoring:
        debug_print("停止 Chrome 監控...")
        existing_monitor.stop_monitoring()
        return None
    else:
        debug_print("啟動 Chrome 監控...")
        monitor = ChromeMonitor()
        monitor.start_monitoring()
        return monitor

# 測試代碼
if __name__ == "__main__":
    monitor = start_chrome_monitor()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.stop_monitoring() 