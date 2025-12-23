# GWC 產規匹配（Web 版）

這是你原本的 CustomTkinter 桌面程式，改成 **Streamlit Web** 的版本。

## 功能
- 來源檔案可多選：每個檔案讀取「最後一個 Sheet」
- 第一列為欄名，第 8 列開始為資料（對應原版）
- 以模板 Sheet1 的 **B 欄 SAP 料號** 做匹配寫入
- 依模板規則檢查：必填 / 選項 / 長度 / 格式（NUM / DATE / CHAR）
- 產出：
  - 主結果檔（可選）
  - 錯誤報表（ErrorLog + SourceCheck）

## 檔案結構
- `app.py`：Streamlit 入口
- `compare_core.py`：核心邏輯（已移除 GUI，改用 BytesIO 下載）
- `requirements.txt`

## 本機執行
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署（Render / Railway / 任何可跑 Python 的平台）
- 只要平台支援 `streamlit run app.py --server.port $PORT --server.address 0.0.0.0` 即可
- 建議使用 Render（Web Service）或 Streamlit Community Cloud
