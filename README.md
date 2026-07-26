# CheckMate — OMR 選擇題閱卷與分析

CheckMate 是適用於 Windows 的 OMR 選擇題閱卷工具。它可從掃描 PDF
辨識答案、找出需要人工覆核的結果、讀取學生資料，並將分數、逐題統計
及課題分析匯出至 Excel。

## 下載

請到 [GitHub Releases](https://github.com/kenkmc/MC_marking/releases/latest)
下載最新版本：

- `CheckMate_Setup_vX.Y.Z.exe`：建議使用的 Windows 安裝程式，不需要管理員權限。
- `CheckMate_vX.Y.Z.zip`：免安裝版本。請先解壓整個 ZIP，再執行
  `CheckMate.exe`；不可只複製 EXE，因為它需要同資料夾內的 `_internal`。
- `SHA256SUMS.txt`：供核對下載檔案完整性。

目前發佈檔未經商業程式碼簽署，因此 Windows SmartScreen 可能顯示提示。

## v1.7 主要功能

- 改良 OMR：支援深色／淺色鉛筆、藍筆、空白及多選答案，不會為模糊結果
  強行猜測。
- 自動校正：固定畫布的歪斜修正，以及可抵抗錯誤定位點的多點旋轉、縮放
  與位移對齊。
- 快速覆核：顯示信心度與異常原因，標記空白、多選、無效或低信心度答案；
  按 `F8` 跳至下一項。
- 題目圖片預覽：直接在程式內檢查及另存答案裁剪圖片。
- 更快的處理：可關閉較慢的文字 OCR；正常辨識不再大量寫入診斷 PNG。
- 便利操作：PDF 拖放、鍵盤快捷鍵，以及自動保存設定。
- 範本與批次處理：範本可按頁面尺寸縮放，並支援共用範本或按檔名配對範本。
- 分析輸出：Excel 成績、總結統計、逐題正確率、課題分析及答案標註圖片。

## 基本使用

1. 匯入 PDF；也可把 PDF 直接拖入視窗。
2. 在第一頁框選選項區、文字區，以及建議使用的兩個或以上定位區。
3. 視需要開啟自動歪斜修正、自動對齊及文字 OCR。
4. 執行全部頁面或目前頁面的辨識。
5. 在右方表格覆核橙色項目；可直接修改答案，或按 `F8` 前往下一項。
6. 匯出 Excel、答案標註圖片，或在需要除錯時匯出診斷資料。

常用快捷鍵：

| 操作 | 快捷鍵 |
|---|---|
| 匯入 PDF | `Ctrl+O` |
| 載入範本 | `Ctrl+L` |
| 儲存範本 | `Ctrl+Shift+S` |
| 辨識全部頁面 | `Ctrl+R` |
| 下一個待覆核項目 | `F8` |

## 提高準確度

- 選項框應緊密覆蓋一整行且間距均勻的選項。
- 掃描時保持相同方向、解像度及良好對比。
- 建立分散在頁面不同位置的定位框，並開啟自動對齊。
- 對淺色、擦改、多選或低信心度項目進行人工覆核。
- 舊版 JSON 範本仍可使用；在 v1.7 重新儲存後會加入頁面尺寸資料，
  方便自動縮放座標。

## 從原始碼執行

需要 Windows 10/11 及 Python 3.10。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

`requirements.txt` 已足夠執行選項 OMR。若要使用 EasyOCR 讀取文字區：

```powershell
python -m pip install torch==2.9.1 torchvision==0.24.1 --index-url https://download.pytorch.org/whl/cpu
python -m pip install -r requirements-ocr.txt
```

EasyOCR 可能會在首次使用時下載辨識模型。PDF 及成績均在本機處理；程式只會
連線檢查 GitHub 更新及按需要下載 OCR 模型。

## 測試與建置

執行核心回歸測試：

```powershell
python -m unittest discover -s tests -v
```

建立完整的 Windows onedir 版本：

```powershell
.\build_exe.ps1
```

輸出位於 `dist\CheckMate\CheckMate.exe`。如已安裝所有建置依賴，可加上
`-SkipInstall`；只需不含 EasyOCR 的較小型本機測試版本，可加上 `-SkipOcr`。

每個版本標籤亦會由 GitHub Actions 執行測試、建立 portable ZIP、建立
per-user Windows 安裝程式、安裝後 smoke test，以及產生 SHA-256 校驗碼。

## English

CheckMate is a Windows OMR marking application that reads answers from scanned PDFs,
highlights uncertain results for review, captures student information, and exports
scores plus question/topic analysis to Excel.

Version 1.7 improves faint-pencil and blue-ink detection, preserves template
coordinates while deskewing, adds robust multi-anchor alignment, confidence-based
review, in-app crop previews, drag-and-drop, shortcuts, scalable templates, and a
faster no-diagnostics-by-default recognition path.

See [CHANGELOG.md](CHANGELOG.md) for details.

## License

[GNU General Public License v3.0](LICENSE)
