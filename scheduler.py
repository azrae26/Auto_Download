import schedule
import time
import threading
from datetime import datetime
from utils import debug_print

class Scheduler:
    def __init__(self, execute_sequence_func):
        self.execute_sequence = execute_sequence_func

    def run_scheduled_task(self):
        """執行排程任務"""
        debug_print(f"[排程器] 開始執行排程任務 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.execute_sequence()

    def start_scheduler(self):
        """啟動排程器"""
        debug_print("[排程器] 排程器已啟動")
        debug_print("[排程器] 已設定每日上午 10:00 及 02:26 自動執行下載任務")
        
        # 設定執行時間
        schedule.every().day.at("10:00").do(self.run_scheduled_task)
        # schedule.every().day.at("02:26").do(self.run_scheduled_task)
        
        while True:
            schedule.run_pending()
            time.sleep(1)

    def init_scheduler(self):
        """初始化排程器"""
        scheduler_thread = threading.Thread(target=self.start_scheduler, daemon=True)
        scheduler_thread.start()
        return scheduler_thread