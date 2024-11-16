from datetime import datetime
import time
import win32gui
import threading
from utils import debug_print
from config import Config

class ChromeMonitor:
    def __init__(self):
        self.chrome_windows = {}
        self.is_monitoring = False
        self.monitor_thread = None
        self.total_pdf_count = 0
        self.total_ppt_count = 0
        self.opened_pdfs = []
        self.last_window_title = None
        
    def normalize_pdf_name(self, name):
        """標準化 PDF 檔名以便比對"""
        # 移除副檔名 .pdf
        if name.endswith('.pdf'):
            name = name[:-4]
            
        # 從最後一個 - 分隔符開始取
        if '-' in name:
            name = name.split('-')[-1].strip()
            
        return name

    def is_same_pdf(self, pdf_name1, pdf_name2):
        """檢查兩個PDF檔名是否為同一個檔案"""
        return self.normalize_pdf_name(pdf_name1) == self.normalize_pdf_name(pdf_name2)
    
    def is_pdf_in_list(self, pdf_name):
        """檢查PDF是否已在清單中（使用模糊比對）"""
        for opened_pdf in self.opened_pdfs:
            if self.is_same_pdf(pdf_name, opened_pdf):
                return True
        return False

    def monitor_loop(self):
        """持續監控檔案視窗"""
        debug_print("\n=== 開始檔案監控 ===")
        
        while self.is_monitoring:
            try:
                current_files = self.find_files()
                
                if current_files:
                    current_window = current_files[0]
                    _, title, file_type, pdf_name = current_window
                    
                    if title != self.last_window_title:
                        self.last_window_title = title
                        
                        if file_type == 'pdf' and pdf_name:
                            # 檢查是否已存在相同的PDF
                            found_match = False
                            for opened_pdf in self.opened_pdfs:
                                if self.is_same_pdf(pdf_name, opened_pdf):
                                    found_match = True
                                    break
                            
                            if not found_match:
                                # 確實是新的PDF
                                debug_print(f"\n新開啟 PDF: {pdf_name}")
                                self.opened_pdfs.append(pdf_name)
                                self.total_pdf_count += 1
                                self._print_status()
                            else:
                                debug_print(f"\n檔案名稱變化: {pdf_name}")
                
                time.sleep(0.2)
                
            except Exception as e:
                debug_print(f"監控過程發生錯誤: {str(e)}")
                time.sleep(0.1)

    def _print_status(self):
        """輸出當前狀態"""
        debug_print("\n=== 當前檔案狀態 ===")
        debug_print(f"PDF 開啟數量: {self.total_pdf_count}")
        if self.opened_pdfs:
            debug_print("已開啟的 PDF 檔案:")
            for pdf in self.opened_pdfs:
                debug_print(f"  - {pdf}")
        else:
            debug_print("目前沒有開啟的 PDF 檔案")
        debug_print(f"PPT 開啟數量: {self.total_ppt_count}")
        debug_print("=====================")

    def start_monitoring(self):
        """開始監控"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
            self.monitor_thread.start()
            debug_print("檔案監控已啟動")
    
    def stop_monitoring(self):
        """停止監控"""
        self.is_monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=1.0)
            debug_print("檔案監控已停止")

    def extract_pdf_name(self, title):
        """從視窗標題中提取 PDF 檔案名稱"""
        title_lower = title.lower()
        if '.pdf' in title_lower:
            # 提取包含 .pdf 的完整檔名
            parts = title_lower.split('.pdf')
            # 從後往前找到最後一個斜線或反斜線的位置
            filename = parts[0].split('\\')[-1].split('/')[-1].strip()
            return filename + '.pdf'
        return None

    def find_files(self):
        """找到所有的 PDF 和 PPT 檔案視窗"""
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                title_lower = title.lower()
                if '.pdf' in title_lower:
                    pdf_name = self.extract_pdf_name(title)
                    if pdf_name:
                        windows.append((hwnd, title, 'pdf', pdf_name))
                elif any(ext in title_lower for ext in ['.ppt', '.pptx']):
                    windows.append((hwnd, title, 'ppt', None))
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

def start_chrome_monitor(existing_monitor=None):
    """啟動或切換檔案監控"""
    if existing_monitor and existing_monitor.is_monitoring:
        debug_print("停止檔案監控...")
        existing_monitor.stop_monitoring()
        return None
    else:
        debug_print("啟動檔案監控...")
        monitor = ChromeMonitor()
        monitor.start_monitoring()
        return monitor
"""""
if __name__ == "__main__":
    monitor = start_chrome_monitor()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.stop_monitoring()
"""""

if __name__ == "__main__":
    monitor = ChromeMonitor()
    monitor.start_monitoring()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.stop_monitoring()