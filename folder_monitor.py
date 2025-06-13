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
        self.target_path = "I:\\共用雲端硬碟\\商拓管理\\券商研究報告分享\\填報告\\test不用填"
        # 新增第二個目標路徑，根據當天日期動態生成
        today_str = datetime.now().strftime("%Y%m%d")
        self.target_path_2 = f"I:\\共用雲端硬碟\\商拓管理\\券商研究報告分享\\每日研究報告任務\\{today_str}"
        self.is_monitoring = False
        self.today_files = []
        self.last_file_count = 0
        self.weekdays = {
            'Monday': '一', 'Tuesday': '二', 'Wednesday': '三',
            'Thursday': '四', 'Friday': '五', 'Saturday': '六', 'Sunday': '日'
        }
        # 排除掃描的檔案
        self.exclude_files = ['sync.ffs_db']
        # 使用正則表達式來匹配要排除的檔案
        self.exclude_pattern = re.compile(
            r'sync\.ffs_db|_公司|晨訊|晨報|日報|週報|月報|週刊|月刊|'
            r'策略|產業|專題|焦點|重要財報|營收檢討|ESG報告|'
            r',jp(?!M)|\.JP(?!M)|-JP(?!M)|_JP(?!M)|'
            r'KS|\.KS|-KS|_KS|'
            r'Weekly|Monthly|I／O|ETF|TPCA|CTBC|Asia |APAC |Global|Greater|SEMICON|'
            r'Telecoms|Financials|Alert|tracker|supply\s*chain',
            re.IGNORECASE)
        # 日誌輸出被排除的檔名
        debug_print(f"排除條件: {self.exclude_pattern.pattern}", color='light_magenta')
        debug_print(f"複製目標路徑1: {self.target_path}", color='light_cyan')
        debug_print(f"複製目標路徑2: {self.target_path_2}", color='light_cyan')

    def log_total_files(self, total_count):
        """輸出今日檔案總數"""
        debug_print("", color='white')
        debug_print("======== 檔案統計 ========", color='light_yellow')
        debug_print(f"     今日檔案總數: {total_count}", color='light_green')
    
    def log_date_statistics(self, date_counts):
        """輸出日期數量統計"""
        sorted_dates = sorted(date_counts.items(), 
                            key=lambda x: (int(x[0].split('/')[0]), int(x[0].split('/')[1])))
        for date, count in sorted_dates:
            month, day = date.split('/')
            weekday = self.weekdays[datetime.strptime(f"2024{month}{day}", "%Y%m%d").strftime('%A')] # 取得星期幾
            debug_print(f"{month} / {day} （{weekday}）：{count} 個檔案", color='light_cyan')
        debug_print("==========================", color='light_yellow')
    
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
            
            debug_print(f"總計收集到 {total_files} 個檔案", color='light_yellow')

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
                debug_print("今日沒有新檔案", color='light_yellow')
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
                    # 格式化列表中的每個檔案名稱
                    normalized_lists[date_name] = [] # 初始化列表
                    for file_name in file_list:
                        if file_name:
                            normalized_name = self._normalize_filename(str(file_name), has_extension=False) # 格式化檔名
                            normalized_lists[date_name].append(normalized_name) # 存入格式化後的檔名
                            # 輸出格式化前後的檔名以便除錯
                            debug_print(f"檔案: 原始={file_name} -> 格式化={normalized_name}", color='white')
                else:
                    normalized_lists[date_name] = []
            
            # 初始化匹配統計
            match_stats = {date_name: 0 for date_name in list_files_dict.keys()}
            matching_results = {}
            
            # 分析每個新檔案
            debug_print("=== 檔案匹配詳情 ===", color='light_cyan')
            for new_file in new_files:
                matches = []
                normalized_new_file = self._normalize_filename(new_file, has_extension=True)
                # 輸出格式化前後的檔名以便除錯
                debug_print(f"檔案: {new_file}", color='white')
                debug_print(f"格式化後: {normalized_new_file}", color='white')
                
                # 比對每個時間點的列表
                for date_name, normalized_file_list in normalized_lists.items():
                    if normalized_new_file in normalized_file_list:
                        matches.append(date_name)
                        match_stats[date_name] += 1
                
                if matches:
                    match_dates = [f"{name} ({date_mapping[name].strftime('%m/%d')})" for name in matches]
                    debug_print(f"匹配: {', '.join(match_dates)}", color='light_magenta')
                else:
                    debug_print("未匹配任何列表", color='light_red')
                
                matching_results[new_file] = matches
            
            # 輸出匹配統計
            debug_print(f"=== {today.strftime('%Y-%m-%d')} 新檔案配統計 ===", color='light_cyan')
            debug_print(f"今日總共有 {len(new_files)} 個新檔案", color='light_yellow')
            for date_name, count in match_stats.items():
                date = date_mapping[date_name]
                weekday = self.weekdays[date.strftime('%A')]
                debug_print(f"{date_name} ({date.strftime('%m/%d')} {weekday}) 新檔案數: {count}", color='light_yellow')
            
            debug_print("========== 分析完成 ==========", color='light_cyan')
            return matching_results

        except Exception as e:
            debug_print(f"分析檔案匹配時發生錯誤: {str(e)}", color='light_red')
            return {}

    def _normalize_filename(self, filename, has_extension=False):
        """
        標準化檔名以便比對，只保留中文、英文、數字
        Args:
            filename: 要處理的檔名
            has_extension: 是否包含副檔名，True=需要移除副檔名，False=不需要移除副檔名
        """
        try:
            # 1. 如果有副檔名且需要移除，則移除最後一個點及其後的所有字元
            name = re.split(r'\.[^.]*$', filename)[0] if has_extension else filename
            
            # 2. 移除所有非中英數字元
            return re.sub(r'[^a-zA-Z0-9\u4e00-\u9fff]+', '', name)
        except Exception as e:
            debug_print(f"標準化檔名時發生錯誤: {str(e)}", color='light_red')
            return filename

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
        debug_print(f"開始監控資料夾: {self.folder_path}")
        self.scan_new_files()
        
    def stop_monitoring(self):
        """停止監控"""
        self.is_monitoring = False
        debug_print("資料夾監控已停止", color='light_yellow')
    
    def copy_file_to_target(self, filename):
        """複製單個檔案到兩個目標資料夾，如果資料夾不存在則建立，包含重試機制"""
        max_retries = 2  # 減少重試次數
        retry_delay = 0.2  # 縮短初始延遲
        
        # 提前檢查來源檔案
        source = os.path.join(self.folder_path, filename)
        if not os.path.exists(source):
            debug_print(f"來源檔案不存在: {filename}", color='light_red')
            return False
        
        # 提前建立目標資料夾
        try:
            if not os.path.exists(self.target_path):
                os.makedirs(self.target_path)
                debug_print(f"已建立目標資料夾1: {self.target_path}", color='light_green')
            
            if not os.path.exists(self.target_path_2):
                os.makedirs(self.target_path_2)
                debug_print(f"已建立目標資料夾2: {self.target_path_2}", color='light_green')
        except Exception as e:
            debug_print(f"建立目標資料夾失敗: {str(e)}", color='light_red')
            return False
        
        # 準備目標路徑
        target1 = os.path.join(self.target_path, filename)
        target2 = os.path.join(self.target_path_2, filename)
        
        for attempt in range(max_retries):
            try:
                # 直接複製，不預先等待
                shutil.copy2(source, target1)
                shutil.copy2(source, target2)
                
                # 只在成功時才顯示詳細訊息
                if attempt > 0:  # 如果是重試成功的
                    debug_print(f"重試成功複製: {filename}", color='light_green')
                
                return True  # 成功複製
                
            except PermissionError as e:
                if "WinError 32" in str(e) or "另一個程序使用" in str(e):
                    if attempt < max_retries - 1:
                        debug_print(f"檔案被佔用，{retry_delay:.1f}秒後重試: {filename}", color='light_yellow')
                        time.sleep(retry_delay)
                        retry_delay *= 1.5  # 較小的延遲增長
                        continue
                    else:
                        debug_print(f"複製失敗 (檔案被佔用): {filename}", color='light_red')
                        return False
                else:
                    debug_print(f"複製失敗 (權限錯誤): {filename}", color='light_red')
                    return False
            except Exception as e:
                if attempt < max_retries - 1:
                    debug_print(f"複製錯誤，重試中: {filename}", color='light_yellow')
                    time.sleep(retry_delay)
                    continue
                else:
                    debug_print(f"複製失敗: {filename}, 錯誤: {str(e)}", color='light_red')
                    return False
        
        return False
    
    def copy_today_files(self):
        """複製今日所有新檔案到兩個位置，但排除特定檔案"""
        today = datetime.now().date()
        copied_files = []  # 成功複製的檔案
        failed_files = []  # 複製失敗的檔案
        excluded_files = []  # 被排除的檔案
        
        try:
            debug_print("開始快速複製檔案...", color='light_cyan')
            
            # 掃描檔案
            all_files = [f for f in os.listdir(self.folder_path) 
                        if os.path.isfile(os.path.join(self.folder_path, f))]
            today_files = []
            
            # 篩選今日檔案
            for file in all_files:
                file_path = os.path.join(self.folder_path, file)
                if datetime.fromtimestamp(os.path.getctime(file_path)).date() == today:
                    today_files.append(file)
            
            if not today_files:
                debug_print("今日沒有新檔案需要複製", color='light_yellow')
                return
            
            debug_print(f"發現 {len(today_files)} 個今日檔案，開始處理...", color='light_blue')
            
            # 處理每個檔案
            for i, file in enumerate(today_files, 1):
                # 使用正則表達式檢查是否為排除的檔案
                if self.exclude_pattern.search(file):
                    match = self.exclude_pattern.search(file)
                    excluded_files.append((file, match.group()))
                    continue
                
                # 顯示進度
                debug_print(f"[{i}/{len(today_files)}] 複製: {file}", color='white')
                
                # 複製檔案到兩個位置
                if self.copy_file_to_target(file):
                    copied_files.append(file)
                else:
                    failed_files.append(file)
            
            # 簡化的結果統計
            debug_print("======= 複製完成 =======", color='light_cyan')
            debug_print(f"✅ 成功: {len(copied_files)} 個檔案", color='light_green')
            
            if failed_files:
                debug_print(f"❌ 失敗: {len(failed_files)} 個檔案", color='light_red')
                for file in failed_files:
                    debug_print(f"   • {file}", color='light_red')
            
            if excluded_files:
                debug_print(f"⏭️  排除: {len(excluded_files)} 個檔案", color='light_yellow')
            
            debug_print(f"📁 目標位置1: {self.target_path}", color='light_blue')
            debug_print(f"📁 目標位置2: {self.target_path_2}", color='light_blue')
            debug_print("========================", color='light_cyan')

        except Exception as e:
            debug_print(f"複製過程發生錯誤: {str(e)}", color='light_red')

def start_folder_monitor(existing_monitor=None):
    """啟動或切換資料夾監控"""
    if existing_monitor and existing_monitor.is_monitoring:
        debug_print("停止資料夾監控...", color='light_yellow')
        existing_monitor.stop_monitoring()
        return None
    else:
        debug_print("啟動資料夾監控...", color='light_yellow')
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