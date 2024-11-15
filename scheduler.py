import schedule
import time
import threading
from datetime import datetime
from utils import debug_print

class Scheduler:
    def __init__(self, task_func):
        self.task_func = task_func
        
    def run_scheduled_task(self, schedule_type=None, schedule_date=None):
        """執行排程任務"""
        current_time = datetime.now()
        
        # 檢查是否為單次執行且日期已過
        if schedule_type == 'once' and schedule_date:
            if current_time.date() > schedule_date.date():
                return schedule.CancelJob
                
        debug_print(f"執行排程任務: {current_time.strftime('%Y-%m-%d %H:%M')}", color='light_blue')
        self.task_func()
        
        # 如果是單次執行，完成後取消任務
        if schedule_type == 'once':
            return schedule.CancelJob
            
    def init_scheduler(self):
        """初始化排程器"""
        schedule.clear()  # 清除所有排程
        
        # 從 Config 獲取排程時間
        from main import Config
        schedule_times = Config.get_schedule_times()
        
        for schedule_info in schedule_times:
            time_str = schedule_info['time']  # 時間字串 (HH:MM)
            
            if schedule_info['type'] == 'daily':
                # 每日執行
                schedule.every().day.at(time_str).do(self.run_scheduled_task)
                debug_print(f"每日排程: {time_str}", color='cyan')
                
            elif schedule_info['type'] == 'weekly':
                # 每週特定日執行
                weekday = schedule_info['weekday'].lower()
                getattr(schedule.every(), weekday).at(time_str).do(self.run_scheduled_task)
                debug_print(f"每週排程: 每週{weekday} {time_str}", color='cyan')
                
            elif schedule_info['type'] == 'once':
                # 單次執行
                schedule_date = datetime.strptime(schedule_info['date'], '%Y-%m-%d')
                schedule.every().day.at(time_str).do(
                    self.run_scheduled_task, 
                    schedule_type='once',
                    schedule_date=schedule_date
                )
                debug_print(f"新增單次排程: {schedule_info['date']} {time_str}", color='cyan')
        
        # 啟動排程執行緒
        thread = threading.Thread(target=self._run_scheduler, daemon=True)
        thread.start()
        return thread
    
    def _run_scheduler(self):
        """運行排程器"""
        while True:
            schedule.run_pending()
            time.sleep(1)