class Config:
    """配置類，集中管理所有配置參數"""
    RETRY_LIMIT = 10  # 向上翻頁次數
    CLICK_BATCH_SIZE = 10  # 批次下載檔案數量
    SLEEP_INTERVAL = 0.1  # 基本等待時間為 0.1 秒
    DOUBLE_CLICK_INTERVAL = 0.1  # 雙擊間隔間
    CLICK_INTERVAL = 0.6  # 連續點擊間隔
    MOUSE_MAX_OFFSET = 100  # 滑鼠最大偏移量
    TARGET_WINDOW = "DostocksBiz"
    PROCESS_NAME = "DostocksBiz.exe"

    @staticmethod
    def get_schedule_times():
        """返回排程時間列表"""
        return [
            {'type': 'daily', 'time': '10:00'},  # 每日固定時間
            {'type': 'once', 'date': '2024-11-15', 'time': '10:50'}  # 單次執行時間
        ] 