import streamlit as st
import pandas as pd
import io

# 設定網頁標題與風格
st.set_page_config(page_title="閱讀獎勵統計系統", page_icon="📚", layout="wide")

# 使用原始字串 (r) 避免語法警告
st.title("📚 學生閱讀獎勵自動統計系統")
st.markdown(r"""
### 運作邏輯：
1. **條件過濾**：僅擷取 **B 欄 (座號)** 有資料，且 **F 欄 (數量)** $\ge 6$ 的列。
2. **欄位擷取**：自動提取 A (班級)、B (座號)、C (姓名)、F (數量)。
3. **獎勵判定**：同一位學生在所有上傳檔案中，若達標次數 **$\ge 3$ 次**，即獲得精美禮物。
""")

# 檔案上傳
uploaded_files = st.file_uploader("請上傳多個 Excel 檔案 (.xlsx, .xls)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    all_rows = []
    
    for uploaded_file in uploaded_files:
        try:
            # 判斷引擎：.xls 使用 xlrd, .xlsx 使用預設引擎
            engine = "xlrd" if uploaded_file.name.endswith(".xls") else None
            df = pd.read_excel(uploaded_file, header=None, engine=engine)
            
            # 檢查檔案欄位數量是否足夠 (至少到 F 欄，即索引 5)
            if df.shape[1] >= 6:
                # 跳過第一列標題
                data_df = df.iloc[1:].copy()
                
                # 數值轉換：將 F 欄轉換為數字，無法轉換的會變成 NaN
                data_df[5] = pd.to_numeric(data_df[5], errors='coerce')
                
                # 核心篩選邏輯：B 欄 (索引 1) 不為空 且 F 欄 (索引 5) >= 6
                mask = (data_df[1].notna()) & (data_df[5] >= 6)
                filtered = data_df[mask].copy()
                
                if not filtered.empty:
                    # 擷取 A, B, C, F 欄位
                    target_columns = filtered[[0, 1, 2, 5]]
                    all_rows.append(target_columns)
            
        except Exception as e:
            st.error(f"檔案 {uploaded_file.name} 處理時發生錯誤: {e}")

    if all_rows:
        # 合併所有檔案的達標紀錄
        combined_df = pd.concat(all_rows)
        combined_df.columns = ["班級", "座號", "姓名", "單次數量"]
        
        # 統計每位學生的達標次數 (以班級+座號+姓名為識別)
        summary = combined_df.groupby(["班級", "座號", "姓名"]).size().reset_index(name="達成總次數")
        
        # 排序：先按班級，再按座號
        summary = summary.sort_values(by=["班級", "座號"])
        
        # 判定獎勵
        summary["獎勵狀態"] = summary["達成總次數"].apply(lambda x: "★ 領取精美禮物" if x >= 3 else "-")
        
        # 顯示統計報表
        st.subheader("📊 統計結果預覽")
        
        # 樣式設定：達標 3 次者顯示淺紅色背景
        def highlight_award(s):
            return ['background-color: #ffcccc' if s['達成總次數'] >= 3 else '' for _ in s]
        
        st.dataframe(summary.style.apply(highlight_award, axis=1), use_container_width=True)
        
        # 提供下載功能
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary.to_excel(writer, index=False, sheet_name='獎勵名單')
        
        st.download_button(
            label="📥 下載獲獎名單 Excel",
            data=output.getvalue(),
            file_name="閱讀獎勵統計結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("目前尚未發現符合條件 (F 欄 $\ge 6$ 且 B 欄有座號) 的資料。")
