import io
import re
import time
import pandas as pd
from datetime import datetime, timedelta
import xlsxwriter

# 來源資料：實際資料從 Excel 第幾列開始（你的來源是第 8 列）
SOURCE_FIRST_DATA_EXCEL_ROW = 8

_X000D_RE = re.compile(r"_x000D_", re.IGNORECASE)

def is_empty(val):
    return (
        val is None
        or (isinstance(val, float) and pd.isna(val))
        or (isinstance(val, str) and val.strip() == "")
    )

def clean_text(val):
    """
    去除 Tab / Enter / CR / Excel XML 的 _x000D_
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None

    s = str(val)

    # 1) 清 Excel XML 的 CR 表示法
    s = _X000D_RE.sub("", s)

    # 2) 清真正的控制字元
    s = s.replace("\t", "").replace("\n", "").replace("\r", "")

    # 3) 其他控制字元
    s = re.sub(r"[\x00-\x1F\x7F]", "", s)

    s = s.strip()
    return s if s else None

def normalize_date(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None

    # Excel serial
    if isinstance(val, (int, float)):
        try:
            iv = int(val)
            if 30000 <= iv <= 60000:
                base = datetime(1899, 12, 30)
                dt = base + timedelta(days=iv)
                return dt.strftime("%Y%m%d")
        except Exception:
            pass

    s = str(val).strip()

    if re.fullmatch(r"\d{8}", s):
        return s

    if re.fullmatch(r"\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}", s):
        y, m, d = re.split(r"[-/\.]", s)
        return f"{y}{m.zfill(2)}{d.zfill(2)}"

    if re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2} \d{2}:\d{2}:\d{2}", s):
        date_part = s.split()[0]
        y, m, d = re.split(r"[-/]", date_part)
        return f"{y}{m.zfill(2)}{d.zfill(2)}"

    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", s):
        a, b, y = s.split("/")
        a_i, b_i = int(a), int(b)
        if a_i > 12:
            d, m = a, b
        elif b_i > 12:
            d, m = b, a
        else:
            d, m = a, b
        return f"{y}{m.zfill(2)}{d.zfill(2)}"

    if re.fullmatch(r"\d{2}/\d{2}/\d{2}", s):
        y, m, d = s.split("/")
        y_i = int(y)
        year = 2000 + y_i if y_i <= 50 else 1900 + y_i
        return f"{year}{m.zfill(2)}{d.zfill(2)}"

    return s

def check_format(value, type_code):
    if value is None:
        return True, None

    s = str(value).strip()
    if s == "":
        return True, None

    t_raw = str(type_code).strip().upper()

    base_type = None
    precision = None
    scale = None

    m = re.search(r"(NUM|NUMBER)\s*\(\s*([0-9]+)\s*[,，\.]\s*([0-9]+)\s*\)", t_raw)
    if m:
        base_type = "NUM"
        precision = int(m.group(2))
        scale = int(m.group(3))
    else:
        if t_raw.startswith("CHAR"):
            base_type = "CHAR"
        elif t_raw.startswith("NUM") or t_raw.startswith("NUMBER"):
            base_type = "NUM"
        elif "DATE" in t_raw:
            base_type = "DATE"
        else:
            base_type = "CHAR"

    if base_type == "CHAR":
        return True, None

    if base_type == "NUM":
        if not re.fullmatch(r"[+-]?[0-9]+(\.[0-9]+)?", s):
            return False, f"格式應為數字(NUM)，實際：{s}"

        num = s.lstrip("+-")
        if "." in num:
            int_part, frac_part = num.split(".", 1)
        else:
            int_part, frac_part = num, ""

        int_digits = len(int_part) if int_part else 0
        frac_digits = len(frac_part)
        total_digits = int_digits + frac_digits

        if precision is not None and total_digits > precision:
            return False, f"數字總位數超過限制：{total_digits}/{precision}（值：{s}）"

        if scale is not None and frac_digits > scale:
            return False, f"小數位數超過限制：{frac_digits}/{scale}（值：{s}）"

        return True, None

    if base_type == "DATE":
        patterns = [
            r"\d{8}",
            r"\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}",
            r"\d{4}[-/]\d{1,2}[-/]\d{1,2} \d{2}:\d{2}:\d{2}",
            r"\d{2}/\d{2}/\d{2}",
            r"\d{1,2}/\d{1,2}/\d{4}",
        ]
        if any(re.fullmatch(p, s) for p in patterns):
            return True, None
        return False, f"格式應為日期(DATE)，實際：{s}"

    return True, None

def to_excel_text(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val)
    s = re.sub(r"[\x00-\x1F\x7F]", "", s)
    return s

def find_source_sap_column(df: pd.DataFrame):
    candidates = []
    for col in df.columns:
        name = str(col).upper()
        if "SAP" in name and ("物料" in name or "料號" in name or "MATERIAL" in name):
            candidates.append(col)
        elif name == "SAP_MATERIAL":
            candidates.append(col)
    return candidates[0] if candidates else None

def get_source_sap_series(df: pd.DataFrame) -> pd.Series:
    col_by_name = find_source_sap_column(df)
    if col_by_name is not None:
        return df[col_by_name]
    if df.shape[1] >= 2:
        return df.iloc[:, 1]
    return pd.Series([None] * len(df), index=df.index)

def _read_excel_bytes(file_bytes: bytes, sheet_name=None, header=None):
    bio = io.BytesIO(file_bytes)
    return pd.read_excel(bio, sheet_name=sheet_name, header=header, engine="openpyxl")

def _read_last_sheet(file_bytes: bytes):
    bio = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(bio, engine="openpyxl")
    last_sheet = xls.sheet_names[-1]
    df = pd.read_excel(bio, sheet_name=last_sheet, header=None, engine="openpyxl")
    return df

def run_core_web(source_files, template_file, only_error_report=False):
    """
    Web 版核心：吃 Streamlit UploadedFile 物件
    回傳 dict: output_bytes/error_bytes/log/stats
    """
    start_time = time.time()
    log_lines = []
    log_lines.append("=== 產規匹配 LOG（Web） ===")
    log_lines.append("[模式] " + ("只輸出錯誤報表" if only_error_report else "完整檢查（主結果 + 錯誤報表）"))

    # -------- 統計 --------
    count_required_err = 0
    count_option_err = 0
    count_length_err = 0
    count_format_err = 0
    count_sap_dup = 0
    source_issue_list = []
    error_cells = {}

    # --------------------------------------------------
    # 0) 讀模板（Sheet1 + Sheet2 選項）
    # --------------------------------------------------
    tpl_bytes = template_file.getvalue()

    options_map = {}
    try:
        opt_df = _read_excel_bytes(tpl_bytes, sheet_name=1, header=None)
        header2 = opt_df.iloc[0]
        for col in range(opt_df.shape[1]):
            field = str(header2[col]).strip()
            if not field:
                continue
            opts = (
                opt_df.iloc[4:44, col]
                .dropna()
                .astype(str)
                .map(clean_text)
                .tolist()
            )
            if opts:
                options_map[field] = set(opts)
    except Exception as e:
        log_lines.append(f"[警告] 第二頁籤讀取失敗：{e}")

    # --------------------------------------------------
    # 1) 合併來源資料：每個來源檔的最後一個 sheet
    # --------------------------------------------------
    merged_df = None
    for uf in source_files:
        df_raw = _read_last_sheet(uf.getvalue())
        header_src = df_raw.iloc[0]
        data = df_raw.iloc[7:].reset_index(drop=True)  # 第 8 列開始
        data.columns = header_src

        if merged_df is None:
            merged_df = data
        else:
            merged_df = pd.concat([merged_df, data], ignore_index=True)

    if merged_df is None or merged_df.empty:
        raise ValueError("來源資料為空，請確認來源檔案內容。")

    # 2) clean
    try:
        merged_clean = merged_df.map(clean_text)  # pandas 2+
    except Exception:
        merged_clean = merged_df.applymap(clean_text)

    # 3) 來源 SAP 欄 & mapping
    source_sap_series_raw = get_source_sap_series(merged_clean)
    source_sap_series = source_sap_series_raw.map(clean_text)

    sap_to_index = {}
    for i, mat in source_sap_series.items():
        if is_empty(mat):
            continue
        if mat not in sap_to_index:
            sap_to_index[mat] = i

    # 4) 讀模板 Sheet1
    target_df = _read_excel_bytes(tpl_bytes, sheet_name=0, header=None)
    header = target_df.iloc[0]
    type_row = target_df.iloc[3].astype(str)
    length_row = target_df.iloc[4]
    require_row = target_df.iloc[5]

    start_row = 6
    SAP_COL_TEMPLATE = 1
    output_df = target_df.copy()

    template_sap_series_raw = target_df.iloc[start_row:, SAP_COL_TEMPLATE]
    template_sap_series = template_sap_series_raw.map(clean_text)

    # 來源 vs 模板：欄位存在性（來源多出來）
    source_columns = {str(c).strip() for c in merged_clean.columns}
    template_columns = {str(h).strip() for h in header}
    missing_cols = sorted(c for c in source_columns if c and c not in template_columns)

    for col_name in missing_cols:
        source_issue_list.append({
            "SourceRow": "-",
            "Material": "-",
            "ErrorType": "來源欄位不存在於模板",
            "Message": f"來源欄位「{col_name}」未出現在模板的欄位列(第1列)中"
        })

    # 來源 vs 模板：料號存在性（來源有、模板沒有）
    template_sap_set = {mat for mat in template_sap_series if not is_empty(mat)}
    for idx, mat in source_sap_series.items():
        if is_empty(mat):
            continue
        if mat not in template_sap_set:
            src_excel_row = SOURCE_FIRST_DATA_EXCEL_ROW + idx
            msg = f"來源料號 {mat} (Row {src_excel_row}) 未在模板 B 欄任一列出現"
            source_issue_list.append({
                "SourceRow": src_excel_row,
                "Material": mat,
                "ErrorType": "來源料號未在模板出現",
                "Message": msg
            })
            log_lines.append(f"[來源料號未在模板出現] {msg}")

    # 模板行 → 來源行
    row_map_template_to_source = {}
    for row_out, mat in template_sap_series.items():
        if is_empty(mat):
            continue
        if mat in sap_to_index:
            row_map_template_to_source[row_out] = sap_to_index[mat]

    # 5) 寫入 + 校驗
    for c in range(len(header)):
        col_name = str(header[c]).strip()
        if not col_name:
            continue
        if col_name not in merged_clean.columns:
            continue

        type_code = type_row[c]
        type_code_str = str(type_code).strip().upper() if not pd.isna(type_code) else ""
        length_limit = length_row[c]
        required_flag = str(require_row[c]).upper() == "V"

        is_date_col = (type_code_str == "DATE" or ("DATE" in str(col_name).upper()))
        series = merged_clean[col_name]

        for row_out, src_idx in row_map_template_to_source.items():
            v_raw = series.iloc[src_idx]
            if isinstance(v_raw, pd.Series):
                v_raw = v_raw.iloc[0]
            v = v_raw

            if is_date_col:
                v = normalize_date(v)

            output_df.iat[row_out, c] = v
            src_excel_row = SOURCE_FIRST_DATA_EXCEL_ROW + src_idx

            # 必填
            if is_empty(v):
                if required_flag:
                    err_type = "必填錯誤"
                    msg = f"Row {src_excel_row} 欄 {col_name} 空白"
                    error_cells.setdefault((row_out, c), []).append((err_type, msg))
                    count_required_err += 1
                    log_lines.append(f"[{err_type}] {msg}")
                continue

            # 選項
            if col_name in options_map and v not in options_map[col_name]:
                err_type = "選項錯誤"
                msg = f"Row {src_excel_row} 欄 {col_name}：{v} 不在允許清單內"
                error_cells.setdefault((row_out, c), []).append((err_type, msg))
                count_option_err += 1
                log_lines.append(f"[{err_type}] {msg}")

            # 長度檢查（保留你的 NUM(整數位,小數位) 與一般長度）
            try:
                if not pd.isna(length_limit):
                    length_spec = str(length_limit).strip()

                    if type_code_str == "NUM":
                        m = re.fullmatch(r"\(\s*(\d+)\s*,\s*(\d+)\s*\)", length_spec)
                        if m:
                            precision = int(m.group(1))
                            scale = int(m.group(2))
                            int_limit = precision - scale

                            s_val = f"{v:.15g}" if isinstance(v, float) else str(v).strip()
                            if not re.fullmatch(r"-?\d+(\.\d+)?", s_val):
                                err_type = "格式錯誤"
                                msg = f"Row {src_excel_row} 欄 {col_name} 應為數字(含小數)，實際：{s_val}"
                                error_cells.setdefault((row_out, c), []).append((err_type, msg))
                                count_format_err += 1
                                log_lines.append(f"[{err_type}] {msg}")
                            else:
                                unsigned = s_val.lstrip("-")
                                if "." in unsigned:
                                    int_part, frac_part = unsigned.split(".", 1)
                                else:
                                    int_part, frac_part = unsigned, ""
                                if len(int_part) > int_limit or len(frac_part) > scale:
                                    err_type = "長度錯誤"
                                    msg = (
                                        f"Row {src_excel_row} 欄 {col_name} 不符合整數 {int_limit} 位、"
                                        f"小數 {scale} 位的限制，值：{s_val}"
                                    )
                                    error_cells.setdefault((row_out, c), []).append((err_type, msg))
                                    count_length_err += 1
                                    log_lines.append(f"[{err_type}] {msg}")
                        else:
                            s_val = f"{v:.15g}" if isinstance(v, float) else str(v).strip()
                            max_len = int(length_spec)
                            raw = s_val.replace("-", "").replace(".", "")
                            if len(raw) > max_len:
                                err_type = "長度錯誤"
                                msg = f"Row {src_excel_row} 欄 {col_name} 實際長度 {len(raw)}/{max_len}，值：{s_val}"
                                error_cells.setdefault((row_out, c), []).append((err_type, msg))
                                count_length_err += 1
                                log_lines.append(f"[{err_type}] {msg}")
                    else:
                        check_str = f"{v:.15g}" if isinstance(v, float) else str(v)
                        max_len = int(length_spec)
                        if len(check_str) > max_len:
                            err_type = "長度錯誤"
                            msg = f"Row {src_excel_row} 欄 {col_name} 實際長度 {len(check_str)}/{max_len}，值：{check_str}"
                            error_cells.setdefault((row_out, c), []).append((err_type, msg))
                            count_length_err += 1
                            log_lines.append(f"[{err_type}] {msg}")
            except Exception:
                pass

            # 格式檢查
            ok, err = check_format(v, type_code_str or type_code)
            if not ok:
                err_type = "格式錯誤"
                msg = f"Row {src_excel_row} 欄 {col_name}：{err}"
                error_cells.setdefault((row_out, c), []).append((err_type, msg))
                count_format_err += 1
                log_lines.append(f"[{err_type}] {msg}")

    # 模板 B 欄 SAP 重複
    sap_clean_template = template_sap_series
    dup_mask = sap_clean_template.duplicated(keep=False) & sap_clean_template.notna()
    dup_indices = sap_clean_template[dup_mask].index
    for row_out in dup_indices:
        mat = sap_clean_template.loc[row_out]
        if row_out in row_map_template_to_source:
            src_idx = row_map_template_to_source[row_out]
            src_excel_row = SOURCE_FIRST_DATA_EXCEL_ROW + src_idx
        else:
            src_excel_row = row_out + 1

        err_type = "SAP重複錯誤"
        msg = f"Row {src_excel_row} 料號 {mat} 重複"
        error_cells.setdefault((row_out, SAP_COL_TEMPLATE), []).append((err_type, msg))
        count_sap_dup += 1
        log_lines.append(f"[{err_type}] {msg}")

    # 成功筆數
    rows_with_error = {r for (r, c) in error_cells.keys()}
    success_row_count = len(row_map_template_to_source) - len(rows_with_error & set(row_map_template_to_source.keys()))

    # --------------------------------------------------
    # 輸出：主結果（可選） + 錯誤報表（有錯才出）
    # --------------------------------------------------
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_bytes = None
    output_name = None
    error_bytes = None
    error_name = None

    # 主結果（用 xlsxwriter，錯誤格紅底黃字）
    if not only_error_report:
        out_bio = io.BytesIO()
        wb = xlsxwriter.Workbook(out_bio, {"in_memory": True})
        ws = wb.add_worksheet("Sheet1")
        err_fmt = wb.add_format({"bg_color": "#FF0000", "font_color": "#FFFF00", "bold": True})

        nrows, ncols = output_df.shape
        for r in range(nrows):
            for c in range(ncols):
                val = output_df.iat[r, c]
                text = to_excel_text("" if pd.isna(val) else val)
                if (r, c) in error_cells:
                    ws.write_string(r, c, text, err_fmt)
                else:
                    ws.write_string(r, c, text)

        wb.close()
        out_bio.seek(0)
        output_bytes = out_bio.getvalue()
        output_name = f"產規匹配結果_{timestamp}.xlsx"

    # 錯誤報表（ErrorLog + SourceCheck）
    has_main_errors = bool(error_cells)
    has_source_errors = bool(source_issue_list)
    if has_main_errors or has_source_errors:
        err_bio = io.BytesIO()
        err_wb = xlsxwriter.Workbook(err_bio, {"in_memory": True})

        if has_main_errors:
            err_ws = err_wb.add_worksheet("ErrorLog")
            headers = ["Row", "Col", "Field", "SAP_Material", "Value", "ErrorType", "ErrorMessage"]
            for i, h in enumerate(headers):
                err_ws.write_string(0, i, h)

            row_idx = 1
            for (row_out, col), msgs in error_cells.items():
                if row_out in row_map_template_to_source:
                    src_idx = row_map_template_to_source[row_out]
                    src_excel_row = SOURCE_FIRST_DATA_EXCEL_ROW + src_idx
                    sap_val = source_sap_series.iloc[src_idx] if 0 <= src_idx < len(source_sap_series) else None
                else:
                    src_excel_row = row_out + 1
                    sap_val = template_sap_series.loc[row_out] if row_out in template_sap_series.index else None

                excel_col = col + 1
                col_name = header[col]
                val = output_df.iat[row_out, col]
                val_text = to_excel_text(val)
                sap_text = to_excel_text(sap_val)

                for err_type, msg in msgs:
                    err_ws.write_string(row_idx, 0, str(src_excel_row))
                    err_ws.write_string(row_idx, 1, str(excel_col))
                    err_ws.write_string(row_idx, 2, to_excel_text(col_name))
                    err_ws.write_string(row_idx, 3, sap_text)
                    err_ws.write_string(row_idx, 4, val_text)
                    err_ws.write_string(row_idx, 5, to_excel_text(err_type))
                    err_ws.write_string(row_idx, 6, to_excel_text(msg))
                    row_idx += 1

        if has_source_errors:
            src_ws = err_wb.add_worksheet("SourceCheck")
            src_headers = ["SourceRow", "SAP_Material", "ErrorType", "Message"]
            for i, h in enumerate(src_headers):
                src_ws.write_string(0, i, h)

            row_idx = 1
            for rec in source_issue_list:
                src_ws.write_string(row_idx, 0, to_excel_text(rec.get("SourceRow", "")))
                src_ws.write_string(row_idx, 1, to_excel_text(rec.get("Material", "")))
                src_ws.write_string(row_idx, 2, to_excel_text(rec.get("ErrorType", "")))
                src_ws.write_string(row_idx, 3, to_excel_text(rec.get("Message", "")))
                row_idx += 1

        err_wb.close()
        err_bio.seek(0)
        error_bytes = err_bio.getvalue()
        error_name = f"產規匹配錯誤報表_{timestamp}.xlsx"

    duration = round(time.time() - start_time, 2)

    stats = {
        "成功轉換列數（有來源且無錯誤）": success_row_count,
        "有錯誤列數": len(rows_with_error),
        "參與匹配的模板列數": len(row_map_template_to_source),
        "必填錯誤": count_required_err,
        "選項錯誤": count_option_err,
        "長度錯誤": count_length_err,
        "格式錯誤": count_format_err,
        "SAP料號重複": count_sap_dup,
        "來源資料檢查錯誤（欄位/料號）": len(source_issue_list),
        "錯誤格數（cell 維度）": len(error_cells),
        "耗時(秒)": duration,
    }

    # LOG 結尾統計
    log_lines.append("")
    log_lines.append("=== 數量統計 ===")
    for k, v in stats.items():
        log_lines.append(f"{k}：{v}")

    return {
        "output_bytes": output_bytes,
        "output_name": output_name,
        "error_bytes": error_bytes,
        "error_name": error_name,
        "log": "\n".join(log_lines),
        "stats": stats,
    }
