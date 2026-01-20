# 發票檔案重新命名工具 (Invoice File Renamer)
這是一個基於 Python 與 PyQt6 開發的桌面應用程式，旨在透過 Excel 表單自動化批次重新命名 PDF 檔案（如發票、收據或掃描文件）。

## 主要功能
- 介面友善：簡潔的圖形化介面，操作直覺。

- 批次處理：一次處理成百上千個檔案，省去手動更改的時間。

- 防錯機制：

    - 自動檢查路徑是否存在。

    - 防止檔名重複覆蓋。

    - 即時顯示處理進度與日誌。

- 欄位對應：彈性對應 Excel 中的「原始掃描檔名」與「目標新檔名」。

## 安裝需求
在使用此工具前，請確保您的環境已安裝 Python 3.8+ 以及相關依賴套件：

```bash
pip install PyQt6 pandas openpyxl
```

## Excel 準備規範
Excel 檔案必須包含以下兩個關鍵欄位（請確保在第一個分頁 Sheet1）：

| 掃描檔名 (包含副檔名) | 新檔名 (不含副檔名) |
| --------------------- | ------------------- |
| SCAN001.pdf           | 20231005_台電電費單  |
| SCAN002.pdf           | 20231010_自來水費    |

注意： 程式會自動為「新檔名」補上 .pdf 副檔名。

## 使用步驟
啟動程式：執行 python your_script_name.py。

選擇 Excel：點擊「瀏覽...」選擇準備好的 Excel 對照表。

選擇目錄：點擊「瀏覽...」選擇存放原始 PDF 檔案的資料夾。

開始執行：確認路徑正確後，按下「開始重新命名」按鈕。

查看結果：處理完成後會跳出視窗告知成功與失敗的數量。

## 注意事項
檔案鎖定：執行時請確保要更名的 PDF 檔案未被其他程式（如 Adobe Reader）開啟。

檔名衝突：如果目標資料夾內已存在相同名稱的檔案，程式會跳過該檔案以保護資料。

路徑檢查：請確保 Excel 中的「掃描檔名」與資料夾內的檔案名稱完全一致（包含大小寫）。

## 技術堆疊
GUI 框架: PyQt6

數據處理: Pandas

檔案操作: Python os 模組

## 如何封裝成 Windows 執行檔 (.exe)
我們建議使用 PyInstaller 來進行封裝。這可以將所有的 Python 依賴（Pandas, PyQt6 等）打包成單一個檔案。

1. 安裝封裝工具

開啟終端機（CMD 或 PowerShell），安裝 PyInstaller：

```bash
pip install pyinstaller
```

2. 執行封裝指令

在您的程式碼目錄下，執行以下指令：

```bash
pyinstaller --noconsole --onefile --name "發票重新命名工具" main.py
```
參數說明：

- `--noconsole`：執行時不會彈出黑色的 CMD 視窗（適合 GUI 程式）。

- `--onefile`：將所有內容打包成一個獨立的 .exe 檔案。

- `--name "發票重新命名工具"`：指定產生的檔名。

- `main.py`：請換成您目前的 Python 檔名。

3. 取得成果

執行完成後，您的目錄下會多出幾個資料夾：

- `dist/`：這最重要！ 您的 .exe 執行檔就在這裡面。

- `build/`：封裝過程產生的暫存檔，打包完可以刪除。

- `filename.spec`：設定檔，下次封裝時可以使用，可保留。

## 進階：加上自訂圖示 (Icon)
如果您有 .ico 格式的圖示檔案，可以使用以下指令讓程式更專業：

```bash
pyinstaller --noconsole --onefile --icon=app_icon.ico --name "發票重新命名工具" main.py
```