"""控件資訊管理模組"""

# 列表控件 ID
LIST_CONTROLS = {
    'morning': {
        'id': 'listBoxMorningReports',
        'name': '晨會報告',
        'type': 'List'
    },
    'research': {
        'id': 'listBoxResearchReports',
        'name': '研究報告',
        'type': 'List'
    },
    'industry': {
        'id': 'listBoxIndustryReports',
        'name': '產業報告',
        'type': 'List'
    }
}

# 日曆控件
CALENDAR_CONTROL = {
    'element_info': {
        'control_type': 'Pane'          # 類型
    },
    'class_name': "WindowsForms10.Window.8.app.0.32f6d92_r8_ad1",  # 類別名稱
    'type_name': 'UIAWrapper',          # 類別
    'size': {
        'width': 185,
        'height': 165
    },
    'layout': {
        'title_height': 52,     # 標題區域高度
        'date_area': {
            'start_y': 52,      # 日期區域起始位置
            'grid_rows': 6,     # 日期網格列數
            'grid_cols': 7      # 日期網格行數
        }
    }
}

# 標籤控件
TAB_CONTROLS = {
    'daily_report': {
        'title': "每日報告",
        'type': "TabItem"
    }
}

# 按鈕控件
BUTTON_CONTROLS = {
    # 可以添加按鈕控件的資訊
}

# 其他控件
OTHER_CONTROLS = {
    # 可以添加其他類型控件的資訊
}

def get_list_control_info(list_type):
    """獲取指定類型列表的控件資訊"""
    return LIST_CONTROLS.get(list_type)

def get_calendar_info():
    """獲取日曆控件資訊"""
    return CALENDAR_CONTROL

def get_tab_info(tab_name):
    """獲取指定標籤的控件資訊"""
    return TAB_CONTROLS.get(tab_name) 