# HSBC Daily Report

每天抓美股收盤價，依使用者的 HSBC 持股 Excel 算市值／損益，用 ReportLab 輸出 PDF。
Windows 桌面程式（tkinter GUI + `--run` 無視窗模式給工作排程器用），PyInstaller 打包、Inno Setup 做安裝程式。

- `app/hsbc_core.py`：讀 xlsx（唯讀）→ yfinance 抓價 → 算損益 → 畫 PDF，不碰 GUI
- `app/hsbc_config.py`：設定檔（`%LOCALAPPDATA%\HSBC Daily Report\config.json`）與工作排程器
- `app/hsbc_app.py`：GUI 與 `--run` 入口；第三方套件延後到 `main()` 才 import，啟動過程寫 `startup.log`
- 編譯只在 CI 做：`.github/workflows/build-windows.yml`，推 `v*` 標籤會自動發佈 Release

## 這台 Windows 機器的開發環境

- Python 3.12（winget 安裝在 `%LOCALAPPDATA%\Programs\Python\Python312`），與 CI 相同版本
- venv 放在 `%LOCALAPPDATA%\venvs\hsbc-daily-report`，**不要放在專案資料夾裡**：
  專案路徑含中文（`OneDrive\文件`），curl_cffi（yfinance 用）讀不到非 ASCII 路徑下的
  certifi `cacert.pem`，會丟 `curl: (77) error adding trust anchors`；而且 venv 也不該被 OneDrive 同步
- 直接跑核心流程（PDF 輸出到別的資料夾，不會動到正式報表）：

  ```powershell
  $env:PYTHONUTF8=1
  & "$env:LOCALAPPDATA\venvs\hsbc-daily-report\Scripts\python.exe" -c "import sys; sys.path.insert(0,'app'); import hsbc_core as c; c.run_report(r'<xlsx 路徑>', r'<輸出資料夾>')"
  ```

- 開 GUI：`& "$env:LOCALAPPDATA\venvs\hsbc-daily-report\Scripts\python.exe" app\hsbc_app.py`

## 注意事項

- **Smart App Control**：這台機器開著（`VerifiedAndReputablePolicyState=1`），會封鎖信譽不足的未簽章 DLL。
  pandas 3.x 的 `_libs\reshape.pyd` 被擋，所以 `requirements.txt` 把 pandas 釘在 2.2.3。
  升級任何含二進位檔的套件前，先在本機確認 import 不會被擋。被擋的檔案可以查
  `Get-WinEvent -LogName "Microsoft-Windows-CodeIntegrity/Operational"`（事件 3077）
- 使用者的 xlsx 與產生的 PDF 是個人財務資料，`.gitignore` 已排除，絕不能 commit
- 程式對 xlsx 全程唯讀，不可改成會寫回的寫法
