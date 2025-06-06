"""
效能優化配置模組
功能：集中管理所有效能相關的配置參數和優化策略
職責：提供可調整的效能參數，支援不同運行模式的切換
依賴：config.py基礎配置
"""

from config import Config

class PerformanceConfig:
    """效能優化配置類"""
    
    # 批次操作模式配置
    BATCH_MODE_ENABLED = True  # 是否啟用批次模式
    BATCH_CACHE_DURATION = 5   # 檔案位置缓存持續時間（秒）
    
    # 檢查頻率優化
    ELEMENT_CHECK_FREQUENCY = 5    # 每N個檔案檢查一次元素（批次模式）
    ERROR_CHECK_FREQUENCY = 10     # 每N個檔案檢查一次錯誤對話框（批次模式）
    WINDOW_FOREGROUND_CHECK_SKIP = True  # 批次模式下是否跳過視窗前景檢查
    
    # 滑鼠移動優化
    BATCH_MOUSE_DURATION = 0.05    # 批次模式下的滑鼠移動時間
    NORMAL_MOUSE_DURATION = 0.1    # 正常模式下的滑鼠移動時間
    
    # 並行處理優化
    ENABLE_PARALLEL_PRELOAD = True  # 是否啟用並行預載入
    PRELOAD_THREAD_COUNT = 2         # 預載入執行緒數量
    
    # 智能跳過策略
    ENABLE_SMART_SKIP = True         # 是否啟用智能跳過
    SKIP_INVISIBLE_FILES = True      # 是否跳過不可見檔案
    MAX_SCROLL_ATTEMPTS = 3          # 最大滾動嘗試次數
    
    # 效能監控
    ENABLE_PERFORMANCE_LOGGING = True  # 是否啟用效能日誌
    LOG_TIMING_DETAILS = False         # 是否記錄詳細時間
    
    @classmethod
    def get_optimized_settings(cls, file_count):
        """根據檔案數量動態調整設定"""
        if file_count > 50:
            # 大量檔案時的激進優化
            return {
                'element_check_freq': 10,
                'error_check_freq': 20,
                'mouse_duration': 0.03,
                'skip_window_checks': True,
                'cache_duration': 10
            }
        elif file_count > 20:
            # 中等數量檔案的平衡優化
            return {
                'element_check_freq': 7,
                'error_check_freq': 15,
                'mouse_duration': 0.04,
                'skip_window_checks': True,
                'cache_duration': 7
            }
        else:
            # 少量檔案時保持穩定性
            return {
                'element_check_freq': 3,
                'error_check_freq': 5,
                'mouse_duration': 0.06,
                'skip_window_checks': False,
                'cache_duration': 5
            }
    
    @classmethod
    def get_timing_analysis(cls):
        """返回時間分析配置"""
        return {
            'measure_preload_time': cls.ENABLE_PERFORMANCE_LOGGING,
            'measure_click_time': cls.LOG_TIMING_DETAILS,
            'measure_validation_time': cls.LOG_TIMING_DETAILS,
            'measure_total_time': cls.ENABLE_PERFORMANCE_LOGGING
        }

# 效能模式常數
class PerformanceMode:
    """效能模式定義"""
    CONSERVATIVE = "conservative"  # 保守模式：穩定性優先
    BALANCED = "balanced"         # 平衡模式：穩定性與速度並重
    AGGRESSIVE = "aggressive"     # 激進模式：速度優先
    
    @classmethod
    def get_mode_config(cls, mode):
        """獲取指定模式的配置"""
        configs = {
            cls.CONSERVATIVE: {
                'batch_mode': False,
                'element_check_freq': 1,
                'error_check_freq': 1,
                'mouse_duration': 0.1,
                'cache_enabled': False
            },
            cls.BALANCED: {
                'batch_mode': True,
                'element_check_freq': 5,
                'error_check_freq': 10,
                'mouse_duration': 0.05,
                'cache_enabled': True
            },
            cls.AGGRESSIVE: {
                'batch_mode': True,
                'element_check_freq': 10,
                'error_check_freq': 20,
                'mouse_duration': 0.03,
                'cache_enabled': True
            }
        }
        return configs.get(mode, configs[cls.BALANCED])

# 時間測量工具
class TimingProfiler:
    """效能計時器"""
    def __init__(self):
        self.timings = {}
        self.enabled = PerformanceConfig.ENABLE_PERFORMANCE_LOGGING
    
    def start_timer(self, operation_name):
        """開始計時"""
        if self.enabled:
            import time
            self.timings[operation_name] = time.time()
    
    def end_timer(self, operation_name):
        """結束計時並返回耗時"""
        if self.enabled and operation_name in self.timings:
            import time
            duration = time.time() - self.timings[operation_name]
            del self.timings[operation_name]
            return duration
        return 0
    
    def log_timing(self, operation_name, duration):
        """記錄計時結果"""
        if self.enabled:
            from utils import debug_print
            debug_print(f"⏱️ {operation_name}: {duration:.3f}秒", color='light_blue')

# 全域計時器實例
profiler = TimingProfiler() 