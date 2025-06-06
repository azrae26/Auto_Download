from utils import debug_print

def test_terminal_support():
    """測試終端機對各種格式的支援情況"""
    debug_print("\n=== 終端機支援測試 ===", color='light_cyan')
    
    # 測試顏色
    for color in ['red', 'light_red', 'green', 'light_green', 'blue', 'light_blue',
                   'yellow', 'light_yellow', 'magenta', 'light_magenta', 'cyan', 'light_cyan']:
        debug_print(f"這是 {color} 顏色的文字", color=color)
    
    # 測試背景色
    debug_print("\n測試背景色:", color='white')
    for bg in ['bg_red', 'bg_green', 'bg_blue', 'bg_yellow', 'bg_magenta', 'bg_cyan']:
        debug_print(f"這是 {bg} 背景的文字", bg_color=bg)
    
    # 測試粗體
    debug_print("\n測試粗體:", color='white')
    debug_print("這是白色", color='white')
    debug_print("這是白色粗體", color='white', bold=True)
    debug_print("這是亮白色", color='light_white')
    debug_print("這是亮白色粗體", color='light_white', bold=True)
    debug_print("這是紅色", color='red')
    debug_print("這是紅色粗體", color='red', bold=True)
    debug_print("這是亮紅色", color='light_red')
    debug_print("這是亮紅色粗體", color='light_red', bold=True)
    debug_print("這是黃色", color='light_yellow')
    debug_print("這是黃色粗體", color='light_yellow', bold=True)
    debug_print("這是亮黃色", color='light_yellow') 
    debug_print("這是亮黃色粗體", color='light_yellow', bold=True)
    debug_print("這是藍色", color='blue')
    debug_print("這是藍色粗體", color='blue', bold=True)
    debug_print("這是亮藍色", color='light_blue')
    debug_print("這是亮藍色粗體", color='light_blue', bold=True)
    

    # 測試組合效果
    debug_print("\n測試組合效果:", color='white')
    debug_print("這是粗體紅色文字", color='red', bold=True)
    debug_print("這是粗體紅色背景文字", color='white', bg_color='bg_red', bold=True)
    
    debug_print("\n=== 測試完成 ===", color='light_cyan')

# 在 main.py 中加入測試選項