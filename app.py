import streamlit as st
from compare_core import run_core_web

st.set_page_config(page_title="GWC 產規匹配（Web 版）", layout="wide")

st.title("GWC 產規明細導入模板校驗產出程式（Web 版）")
st.caption("來源檔案合併 ➜ 產規規則檢查 ➜ 產出錯誤報表 / 匹配結果檔")

with st.sidebar:
    st.header("操作")
    mode = st.radio("模式", ["完整檢查（主結果 + 錯誤報表）", "只輸出錯誤報表"], index=0)
    only_error = (mode == "只輸出錯誤報表")
    st.divider()
    st.markdown("**注意事項**")
    st.markdown(
        "1) 來源檔案：讀取每個檔案的「最後一個 Sheet」\n"
        "2) 第一列為欄名，第 8 列開始為資料（可調整）\n"
        "3) 模板：第 1 列為欄位名，第 4/5/6 列分別為型態/長度/必填\n"
        "4) 第二頁籤（Sheet2）可放選項清單（第 1 列欄名、第 5~44 列選項）"
    )

col1, col2 = st.columns(2)

with col1:
    src_files = st.file_uploader(
        "STEP1：選擇來源資料（可多選）",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

with col2:
    tpl_file = st.file_uploader(
        "STEP2：選擇目標模板",
        type=["xlsx", "xls"],
        accept_multiple_files=False
    )

st.markdown("### 執行")
run = st.button("開始執行", type="primary", use_container_width=True, disabled=(not src_files or not tpl_file))

if run:
    with st.spinner("處理中..."):
        result = run_core_web(
            source_files=src_files,
            template_file=tpl_file,
            only_error_report=only_error
        )

    st.success("完成！")

    # ---- 統計 ----
    s = result["stats"]
    st.subheader("統計")
    st.json(s)

    # ---- 下載 ----
    st.subheader("輸出檔案")
    if result.get("output_bytes") is not None:
        st.download_button(
            "下載主結果檔",
            data=result["output_bytes"],
            file_name=result["output_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("本次未產生主結果檔（你選了只輸出錯誤報表，或沒有需要輸出）。")

    if result.get("error_bytes") is not None:
        st.download_button(
            "下載錯誤報表",
            data=result["error_bytes"],
            file_name=result["error_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("本次未產生錯誤報表（可能沒有任何錯誤）。")

    # ---- LOG ----
    with st.expander("查看 LOG（文字）", expanded=False):
        st.text(result["log"])
