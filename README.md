# DostocksBiz 自動下載系統

## 專案概述

DostocksBiz 自動下載系統是一個針對 DostocksBiz 應用程式的自動化工具，能夠自動下載晨會報告、研究報告和產業報告。系統採用 Python 開發，使用 Windows 自動化技術實現智能文件處理。

**更新時間：** 2025-06-06 22:39:50

## 主要功能

### 📋 核心功能
- **自動檔案下載**：支援晨會報告、研究報告、產業報告的批量下載
- **智能視窗管理**：自動檢測和控制 DostocksBiz 視窗
- **熱鍵快速操作**：提供豐富的快捷鍵操作
- **檔案監控**：即時監控下載資料夾變化
- **Chrome 瀏覽器整合**：監控瀏覽器下載狀態

### 🕒 排程功能
- **定時執行**：支援每日固定時間自動執行
- **單次排程**：支援指定日期時間的一次性執行
- **彈性配置**：可自定義排程時間和執行頻率

### 🖱️ 使用者介面
- **熱鍵系統**：
  - `F1`：下載當前列表所有檔案
  - `F2`：執行完整下載序列
  - `F3`：切換熱鍵開關
  - `F4`：列出所有報告
  - `F5`：開始/停止刷新檢測
  - `F6`：監控 Chrome 瀏覽器
  - `F7`：複製今日檔案
  - `F8`：收集並分析列表
  - `ESC`：緊急停止所有操作

## 系統架構

### 📁 檔案結構
```
Auto_Download/
├── main.py              # 主程式進入點
├── utils.py              # 核心工具函數
├── config.py             # 系統配置參數
├── folder_monitor.py     # 資料夾監控功能
├── chrome_monitor.py     # Chrome 瀏覽器監控
├── calendar_checker.py   # 行事曆檢查功能
├── scheduler.py          # 排程管理
├── get_list_area.py      # 列表區域處理
├── font_size_setter.py   # 字體大小設定
├── control_info.py       # 控制項資訊
├── test_terminal.py      # 終端測試功能
└── requirements.txt      # 專案依賴
```

### 🔧 主要模組

#### 1. FileProcessor 類別
- **功能**：處理文件相關操作
- **職責**：批量下載、視窗管理、檔案狀態追蹤
- **依賴**：pywinauto、win32api、pyautogui

#### 2. MainApp 類別
- **功能**：主應用程式控制器
- **職責**：熱鍵註冊、UI 互動、排程管理
- **依賴**：keyboard、threading、queue

#### 3. 工具函數庫 (utils.py)
- **功能**：核心操作函數集合
- **職責**：視窗操作、滑鼠控制、檔案檢測
- **依賴**：win32gui、pyautogui、pywinauto

## 安裝與設置

### 系統需求
- **作業系統**：Windows 10/11
- **Python 版本**：3.7+
- **目標應用**：DostocksBiz 軟體

### 安裝步驟

1. **安裝 Python 依賴**
   ```bash
   pip install -r requirements.txt
   ```

2. **確認 DostocksBiz 已安裝並可正常運行**

3. **配置系統參數**
   ```python
   # 在 config.py 中調整參數
   CLICK_BATCH_SIZE = 20      # 批次下載檔案數量
   SLEEP_INTERVAL = 0.05      # 基本等待時間
   DOWNLOAD_INTERVAL = 0.01   # 下載間隔
   ```

## 使用說明

### 基本操作

1. **啟動程式**
   ```bash
   python main.py
   ```

2. **選擇 DostocksBiz 視窗**
   - 程式啟動後會列出可用視窗
   - 輸入對應數字選擇目標視窗

3. **開始自動下載**
   - 按 `F1` 下載當前列表
   - 按 `F2` 執行完整序列
   - 按 `ESC` 緊急停止

### 進階功能

#### 排程設置
```python
# 在 config.py 中設定排程時間
def get_schedule_times():
    return [
        {'type': 'daily', 'time': '09:45'},              # 每日 09:45
        {'type': 'once', 'date': '2024-11-15', 'time': '10:50'}  # 單次執行
    ]
```

#### 監控功能
- **資料夾監控**：自動檢測下載完成的檔案
- **Chrome 監控**：追蹤瀏覽器下載狀態
- **視窗監控**：確保目標視窗保持可見狀態

## 效能優化

### 當前效能特色
- **預加載機制**：提前載入列表資訊，減少重複查詢
- **批次處理**：每 20 個檔案關閉一次視窗，避免記憶體累積
- **智能滾動**：自動調整檔案可見性，提高點擊準確度
- **錯誤恢復**：自動檢測和處理各種異常情況

### 配置參數說明
```python
RETRY_LIMIT = 10               # 重試次數限制
CLICK_BATCH_SIZE = 20          # 批次下載數量
SLEEP_INTERVAL = 0.05          # 基本等待時間
DOUBLE_CLICK_INTERVAL = 0.05   # 雙擊間隔
DOWNLOAD_INTERVAL = 0.01       # 下載間隔
CLOSE_WINDOW_INTERVAL = 0.1    # 關閉視窗間隔
```

## 故障排除

### 常見問題

1. **找不到 DostocksBiz 視窗**
   - 確認 DostocksBiz 程式已啟動
   - 檢查視窗標題是否包含 "DostocksBiz"

2. **點擊位置不準確**
   - 檢查螢幕解析度設定
   - 確認 DostocksBiz 視窗未被其他視窗遮擋

3. **下載失敗**
   - 確認網路連線正常
   - 檢查下載資料夾權限
   - 查看錯誤對話框提示

### 調試模式
程式內建詳細的調試輸出，包含：
- 彩色狀態訊息
- 操作時間戳記
- 詳細錯誤資訊
- 檔案處理進度

## 開發資訊

### 技術棧
- **GUI 自動化**：pywinauto, pyautogui
- **系統操作**：win32api, win32gui
- **熱鍵處理**：keyboard
- **檔案監控**：psutil
- **多執行緒**：threading, queue

### 設計原則
- **模組化架構**：功能分離，便於維護
- **錯誤處理**：完善的異常捕獲和恢復機制
- **效能優化**：批次處理和智能快取
- **用戶友好**：直觀的熱鍵操作和狀態提示

### 程式碼風格
- 遵循 PEP 8 規範
- 詳細的中文註解
- 模組化設計
- 異常安全處理

## 授權資訊

此專案為內部使用工具，請遵循公司相關規定使用。

## 更新記錄

- **2025-06-06**：建立 README 文件，系統功能完善
- 支援多列表檔案預加載
- 優化點擊效能和錯誤處理
- 完善熱鍵系統和使用者介面

---

**注意事項**：
- 請確保在使用前備份重要資料
- 建議在測試環境中先行驗證功能
- 遇到問題請查看控制台輸出的詳細資訊 