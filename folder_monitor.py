import os
from datetime import datetime, timedelta
import time
from config import Config, COLORS
from utils import debug_print
import shutil
import re

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
            '_公司', '晨訊', '週報', '策略', '產業', 'ETF', 'supply chain'
        ]
        
    def log_total_files(self, total_count):
        """輸出今日檔案總數"""
        debug_print("", color='white')
        debug_print("======== 檔案統計 ========", color='yellow')
        debug_print(f"     今日檔案總數: {total_count}", color='light_green')
    
    def log_date_statistics(self, date_counts):
        """輸出日期數量統計"""
        sorted_dates = sorted(date_counts.items(), 
                            key=lambda x: (int(x[0].split('/')[0]), int(x[0].split('/')[1])))
        for date, count in sorted_dates:
            month, day = date.split('/')
            weekday = self.weekdays[datetime.strptime(f"2024{month}{day}", "%Y%m%d").strftime('%A')] # 取得星期幾
            debug_print(f"{month} / {day} （{weekday}）：{count} 個檔案", color='light_cyan')
        debug_print("==========================", color='yellow')
    
    def log_new_file(self, filename):
        """輸出新發現的檔案"""
        debug_print(f"發現新檔案: {filename}", color='light_green')

    def scan_new_files(self):
        """掃描今日新檔案"""
        today = datetime.now().date()
        new_files, date_counts = self.scan_files_for_date(today)
        return new_files, date_counts
    
    def scan_files_for_date(self, target_date):
        """掃描指定日期的檔案"""
        files = []
        date_counts = {}
        
        try:
            # 掃描資料夾
            for file in os.listdir(self.folder_path):
                if file in self.exclude_files:
                    continue
                # 取得檔案路徑
                file_path = os.path.join(self.folder_path, file)
                if not os.path.isfile(file_path):
                    continue
                # 取得檔案建立時間
                if datetime.fromtimestamp(os.path.getctime(file_path)).date() == target_date:
                    files.append(file)
                    
                    # 分析日期
                    for part in file.split('_'):
                        if len(part) == 7 and part.startswith('11'):
                            date = f"{part[-4:-2]}/{part[-2:]}"
                            date_counts[date] = date_counts.get(date, 0) + 1
                        
        except Exception as e:
            debug_print(f"掃描檔案時發生錯誤: {str(e)}", color='light_red')
        
        return files, date_counts

    def store_and_analyze_lists(self, today_list=None, yesterday_list=None, last_week_list=None, 
                                last_2week_list=None, last_3week_list=None, last_4week_list=None):
        """存儲並分析各時間點的檔案列表"""
        try:
            # 儲存各時間點的列表
            lists_dict = {
                '今日': today_list or [],
                '昨日': yesterday_list or [],
                '1週前': last_week_list or [],
                '2週前': last_2week_list or [],
                '4週前': last_3week_list or [],
                '8週前': last_4week_list or []
            }

            # 分析列表內容
            debug_print("=== 各時間點列表分析 ===", color='light_cyan')
            total_files = 0
            for date_name, file_list in lists_dict.items():
                file_count = len(file_list) if file_list else 0
                total_files += file_count
                debug_print(f"[{date_name}] 列表檔案數: {file_count}", color='light_blue', bold=True)
                for file in file_list:
                    if file:  # 確保檔案名稱不是空的
                        debug_print(f"- {file}", color='white')
            
            debug_print(f"總計收集到 {total_files} 個檔案", color='yellow')

            # 分析今日新檔案的匹配情況
            if total_files > 0:
                matching_results = self.analyze_new_files_with_lists(lists_dict)
                return matching_results
            else:
                debug_print("沒有收集到任何檔案，跳過匹配分析", color='light_red')
                return {}

        except Exception as e:
            debug_print(f"分析列表時發生錯誤: {str(e)}", color='light_red')
            return {}


    def analyze_new_files_with_lists(self, list_files_dict):
        """分析今日新檔案與各時間點列表的匹配情況"""
        try:
            today = datetime.now().date()
            new_files, _ = self.scan_files_for_date(today)
            
            if not new_files:
                debug_print("今日沒有新檔案", color='yellow')
                return {}
            
            # 定義各時間點的日期
            date_mapping = {
                '今日': today,
                '昨日': today - timedelta(days=1),
                '1週前': today - timedelta(days=7),
                '2週前': today - timedelta(days=14),
                '4週前': today - timedelta(days=28),
                '8週前': today - timedelta(days=56)
            }
            
            debug_print("=== 檔案數列表 ===", color='light_cyan')
            
            # 預先處理列表檔案名稱
            normalized_lists = {}
            # 遍歷所有時間點，包括沒有檔案的
            for date_name, date in date_mapping.items():
                file_list = list_files_dict.get(date_name, [])
                weekday = self.weekdays[date.strftime('%A')]
                debug_print(f"{date_name} ({date.strftime('%m/%d')} {weekday}) 列表檔案數: {len(file_list)}", color='light_blue', bold=True)
                
                if file_list:
                    normalized_lists[date_name] = [
                        self._normalize_filename(str(file_name)) 
                        for file_name in file_list 
                        if file_name
                    ]
                else:
                    # 如果時間點沒有檔案，則設為空列表
                    normalized_lists[date_name] = []
            
            # 初始化匹配統計
            match_stats = {date_name: 0 for date_name in list_files_dict.keys()}
            matching_results = {}
            
            # 分析每個新檔案
            debug_print("=== 檔案匹配詳情 ===", color='light_cyan')
            for new_file in new_files:
                matches = []
                normalized_new_file = self._normalize_filename(new_file)
                
                # 比對每個時間點的列表
                for date_name, normalized_file_list in normalized_lists.items():
                    if normalized_new_file in normalized_file_list:
                        matches.append(date_name)
                        match_stats[date_name] += 1
                
                # 輸出個別檔案的匹配結果
                debug_print(f"檔案: {new_file}", color='white')
                if matches:
                    match_dates = [f"{name} ({date_mapping[name].strftime('%m/%d')})" for name in matches]
                    debug_print(f"匹配: {', '.join(match_dates)}", color='light_magenta')
                else:
                    debug_print("未匹配任何列表", color='light_red')
                
                matching_results[new_file] = matches
            
            # 輸出匹配統計
            debug_print("=== 新檔案配統計 ===", color='light_cyan')
            debug_print(f"今日總共有 {len(new_files)} 個新檔案", color='yellow')
            for date_name, count in match_stats.items():
                date = date_mapping[date_name]
                weekday = self.weekdays[date.strftime('%A')]
                debug_print(f"{date_name} ({date.strftime('%m/%d')} {weekday}) 新檔案數: {count}", color='yellow')
            
            debug_print("=== 分析完成 ===", color='light_cyan')
            return matching_results

        except Exception as e:
            debug_print(f"分析檔案匹配時發生錯誤: {str(e)}", color='light_red')
            return {}

    def _normalize_filename(self, filename):
        """標準化檔名以便比對"""
        # 移除副檔名
        if '.' in filename:
            filename = filename.rsplit('.', 1)[0]
        
        # 移除路徑（如果有）
        filename = filename.split('\\')[-1]
        
        # 移除所有特殊符號和空白
        # 保留中文、英文、數字
        filename = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', filename)
        
        return filename.lower()  # 轉為小寫以忽略大小寫差異

    def scan_new_files_and_log(self):
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
            debug_print(f"掃描資料夾時發生錯誤: {str(e)}", color='light_red')
            
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
            debug_print(f"已複製: {filename}", color='white')
        except Exception as e:
            debug_print(f"複製檔案失敗: {filename}, 錯誤: {str(e)}", color='light_red')
    
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
            debug_print(f"複製檔案過程發生錯誤: {str(e)}", color='light_red')

def start_folder_monitor(existing_monitor=None):
    """啟動或切換資料夾監控"""
    if existing_monitor and existing_monitor.is_monitoring:
        debug_print("停止資料夾監控...", color='yellow')
        existing_monitor.stop_monitoring()
        return None
    else:
        debug_print("啟動資料夾監控...", color='yellow')
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