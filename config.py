from colorama import init, Fore, Back, Style

# 初始化 colorama
init()

# 顏色映射
COLORS = {
    # 文字顏色
    'red': Fore.RED,                      # 錯誤訊息，紅色
    'orange': Fore.LIGHTRED_EX,           # 橘色
    'green': Fore.GREEN,                  # 成功訊息，綠色
    'yellow': Fore.YELLOW,                # 操作訊息，黃色
    'blue': Fore.BLUE,                    # 一般資訊，藍色
    'magenta': Fore.MAGENTA,              # 檢測訊息，洋紅色
    'cyan': Fore.CYAN,                    # 標題，青色
    'light_red': Fore.LIGHTRED_EX,        # 亮紅色
    'light_green': Fore.LIGHTGREEN_EX,    # 亮綠色
    'light_yellow': Fore.LIGHTYELLOW_EX,  # 亮黃色
    'light_blue': Fore.LIGHTBLUE_EX,      # 亮藍色
    'light_magenta': Fore.LIGHTMAGENTA_EX,# 亮洋紅色
    'light_cyan': Fore.LIGHTCYAN_EX,      # 亮青色
    'light_white': Fore.LIGHTWHITE_EX,    # 亮白色
    'dark_grey': Fore.LIGHTBLACK_EX,      # 深灰色
    'reset': Style.RESET_ALL,             # 重置顏色
    # 背景顏色
    'bg_red': Back.RED,                   # 紅色背景
    'bg_green': Back.GREEN,               # 綠色背景
    'bg_yellow': Back.YELLOW,             # 黃色背景
    'bg_blue': Back.BLUE,                 # 藍色背景
    'bg_magenta': Back.MAGENTA,           # 洋紅色背景
    'bg_cyan': Back.CYAN,                 # 青色背景
    'bg_white': Back.WHITE,               # 白色背景
    'bg_reset': Back.RESET,               # 重置背景
    # 樣式
    'bold': Style.BRIGHT,                  # 粗體
    'dim': Style.DIM,                      # 暗淡
    'normal': Style.NORMAL,                # 正常
    'reset': Style.RESET_ALL,              # 重置樣式
}

class Config:
    """配置類，集中管理所有配置參數"""
    RETRY_LIMIT = 10  # 向上翻頁次數
    CLICK_BATCH_SIZE = 10  # 批次下載檔案數量
    SLEEP_INTERVAL = 0.01  # 基本等待時間
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