import pandas as pd
import os
from tkinter import Tk, filedialog
from datetime import datetime
import traceback

# 預設檔名（如檔名不同你可以自己改這兩行）
DEFAULT_SOURCE = "GWC_期初上線物料主檔及產規收集_V4.0.xlsx"
DEFAULT_TARGET = "產規批導模板.xlsx"

def select_file(title):
    """開啟選檔視窗"""
    Tk().withdraw()
    return filedialog.askopenfilename(title=title)

def load_file_or_select(default_name, description):
    """
    如找到預設檔案則直接使用，否則跳窗讓使用者選取
    description 例：'來源檔（產規收集模版1.4）'
    """
    if os.path.exists(default_name):
        print(f"✔ 已找到{description}：{default_name}")
        return default_name
    else:
        print(f"⚠ 未找到{description}：{default_name}")
        print(f"→ 請手動選擇{description}")
        return select_file(f"請選擇 {description}")

def main():
    try:
        print("=== 產規檔案匹配程式啟動 ===")

        # 1. 取得來源檔與批導模板檔
        source_file = load_file_or_select(DEFAULT_SOURCE, "來源檔（產規收集模版1.4）")
        target_file = load_file_or_select(DEFAULT_TARGET, "批導模板檔")

        if not source_file or not target_file:
            print("❌ 沒有選擇完整的檔案，程式結束。")
            return

        print("📌 正在讀取來源檔案...")
        # 不指定 header，全部當一般資料，再自行取第 1 列當 header
        source_df = pd.read_excel(source_file, sheet_name="產規收集模版1.4", header=None)

        # 英文欄位在第 1 列（index = 0）
        source_header = source_df.iloc[0]

        # 資料從第 8 列開始（也就是 index = 7）
        source_data = source_df.iloc[7:].reset_index(drop=True)

        print("📌 正在讀取批導模板檔...")
        target_df = pd.read_excel(target_file, header=None)

        # 目標模板的英文欄位也在第 1 列（index = 0）
        target_header = target_df.iloc[0]

        # 要貼入資料的起始列 = 第 7 列（index = 6）
        start_row = 6

        # 複製一份模板
        new_target = target_df.copy()

        # 若目標列數不夠，先擴充
        rows_needed = start_row + len(source_data)
        if len(new_target) < rows_needed:
            extra_rows = rows_needed - len(new_target)
            new_target = pd.concat(
                [new_target, pd.DataFrame([[None] * new_target.shape[1]] * extra_rows)],
                ignore_index=True
            )

        print("📌 正在比對欄位（依英文欄位名稱）...")

        # 逐欄比對：目標欄位名稱 vs 來源欄位名稱（都看第 1 列）
        for col_target in range(len(target_header)):
            target_col_name = str(target_header[col_target]).strip()

            if not target_col_name or target_col_name == "nan":
                continue

            # 在來源 header 裡找欄位名稱一樣的
            match_cols = source_header[source_header == target_col_name].index.tolist()

            if not match_cols:
                # 找不到對應欄位就跳過
                continue

            source_col = match_cols[0]

            # 把來源的資料列（從第 8 列開始）貼到目標（第 7 列開始）
            new_target.iloc[start_row:start_row + len(source_data), col_target] = \
                source_data.iloc[:, source_col].values

        # 產生輸出檔名：產規匹配結果_YYYYMMDD_HHMMSS.xlsx
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"產規匹配結果_{timestamp}.xlsx"

        new_target.to_excel(output_file, header=False, index=False)

        print("✔ 匹配完成！")
        print(f"已產生輸出檔案：{output_file}")
        print("=== 程式執行完畢 ===")

    except Exception as e:
        print("❌ 發生錯誤！")
        print(str(e))
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
