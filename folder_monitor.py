import os
from datetime import datetime
import time
from utils import debug_print
import shutil

class FolderMonitor:
    def __init__(self, folder_path="C:\\temp"):
        self.folder_path = folder_path
        self.target_path = "I:\\共用雲端硬碟\\商拓管理\\券商研究報告分享\\填報告"
        self.is_monitoring = False
        self.today_files = []
        self.last_file_count = 0
        self.weekdays = {
            'Monday': '一', 'Tuesday': '二', 'Wednesday': '三',
            'Thursday': '四', 'Friday': '五', 'Saturday': '六', 'Sunday': '日'
        }
        # 排除掃描的檔案
        self.exclude_files = ['sync.ffs_db']
        # 包含以下關鍵字排除複製
        self.exclude_copy_files = [
            'sync.ffs_db',
            '_公司', '_晨訊', '_週報', '_策略', '_產業', '_ETF', '_supply chain'
        ]
        
    def log_total_files(self, total_count):
        """輸出今日檔案總數"""
        debug_print("", color='white')
        debug_print("======== 檔案統計 ========", color='yellow')
        debug_print(f"     今日檔案總數: {total_count}", color='green')
    
    def log_date_statistics(self, date_counts):
        """輸出日期數量統計"""
        sorted_dates = sorted(date_counts.items(), 
                            key=lambda x: (int(x[0].split('/')[0]), int(x[0].split('/')[1])))
        for date, count in sorted_dates:
            month, day = date.split('/')
            weekday = self.weekdays[datetime.strptime(f"2024{month}{day}", "%Y%m%d").strftime('%A')] # 取得星期幾
            debug_print(f"{month} / {day} （{weekday}）：{count} 個檔案", color='cyan')
        debug_print("==========================", color='yellow')
    
    def log_new_file(self, filename):
        """輸出新發現的檔案"""
        debug_print(f"發現新檔案: {filename}", color='green')
        
    def scan_new_files(self):
        """掃描今日新檔案，並在啟動及檔案數量變化時輸出統計"""
        today = datetime.now().date()
        new_files = []
        date_counts = {}
        
        try:
            # 掃描檔案
            for file in os.listdir(self.folder_path):
                if file in ['sync.ffs_db']:
                    continue
                # 取得檔案路徑
                file_path = os.path.join(self.folder_path, file)
                if not os.path.isfile(file_path):
                    continue
                # 取得檔案建立時間
                if datetime.fromtimestamp(os.path.getctime(file_path)).date() == today:
                    new_files.append(file)
                    if file not in self.today_files:
                        self.today_files.append(file)
                        self.log_new_file(file)
                        
                        # 分析日期
                        for part in file.split('_'):
                            if len(part) == 7 and part.startswith('11'):
                                date = f"{part[-4:-2]}/{part[-2:]}"
                                date_counts[date] = date_counts.get(date, 0) + 1
            
            # 只在檔案數有變化時輸出統計
            current_count = len(new_files)
            if current_count != self.last_file_count:
                self.last_file_count = current_count
                if new_files:
                    self.log_total_files(current_count)
                    self.log_date_statistics(date_counts)
                    
        except Exception as e:
            debug_print(f"掃描資料夾時發生錯誤: {str(e)}")
            
        return new_files
    
    def start_monitoring(self):
        """開始監控"""
        self.is_monitoring = True
        debug_print(f"\n開始監控資料夾: {self.folder_path}")
        self.scan_new_files()
        
    def stop_monitoring(self):
        """停止監控"""
        self.is_monitoring = False
        debug_print("資料夾監控已停止")
    
    def copy_file_to_target(self, filename):
        """複製單個檔案到目標資料夾"""
        try:
            # 取得檔案路徑
            source = os.path.join(self.folder_path, filename)
            # 取得目標路徑
            target = os.path.join(self.target_path, filename)
            # 複製檔案
            shutil.copy2(source, target)
            debug_print(f"已複製: {filename}", color='blue')
        except Exception as e:
            debug_print(f"複製檔案失敗: {filename}, 錯誤: {str(e)}", color='red')
    
    def copy_today_files(self):
        """複製今日所有新檔案，但排除特定檔案"""
        today = datetime.now().date()
        copied_count = 0  # 計數器
        
        try:
            # 掃描檔案
            for file in os.listdir(self.folder_path):
                # 檢查是否為排除的檔案
                if any(exclude in file for exclude in self.exclude_copy_files):
                    continue
                
                # 取得檔案路徑
                file_path = os.path.join(self.folder_path, file)
                if not os.path.isfile(file_path):
                    continue
                
                # 檢查檔案日期
                if datetime.fromtimestamp(os.path.getctime(file_path)).date() == today:
                    self.copy_file_to_target(file)  # 複製檔案
                    copied_count += 1  # 計數加1
            
            # 輸出複製統計
            debug_print("")
            debug_print("======== 複製統計 ========")
            debug_print(f"     成功複製: {copied_count} 個檔案")
            debug_print("==========================")

        except Exception as e:
            debug_print(f"複製檔案過程發生錯誤: {str(e)}")

def start_folder_monitor(existing_monitor=None):
    """啟動或切換資料夾監控"""
    if existing_monitor and existing_monitor.is_monitoring:
        debug_print("停止資料夾監控...")
        existing_monitor.stop_monitoring()
        return None
    else:
        debug_print("啟動資料夾監控...")
        monitor = FolderMonitor()
        monitor.start_monitoring()
        return monitor

if __name__ == "__main__":
    monitor = FolderMonitor()
    monitor.start_monitoring()
    try:
        while True:
            monitor.scan_new_files()
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.stop_monitoring() 