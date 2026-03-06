import streamlit as st
import pandas as pd
import io

# 設定網頁標題
st.set_page_config(page_title="學生閱讀獎勵統計系統", layout="wide")

st.title("📚 學生閱讀獎勵自動統計系統")
st.markdown("""
**規則說明：**
1. 檢查 **B 欄 (座號)** 是否有資料。
2. 檢查 **F 欄 (數量)** 是否 $\ge 6$。
3. 統計每位學生在所有檔案中符合上述條件的**次數**。
4. 累計 **3 次** 以上者標註為「可領取獎品」。
""")

# 檔案上傳元件 (支援多檔上傳)
uploaded_files = st.file_uploader("請選擇多個 Excel 檔案", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for uploaded_file in uploaded_files:
        try:
            # 讀取 Excel，不預設表頭以精確控制欄位索引
            df = pd.read_excel(uploaded_file, header=None)
            
            # 假設第 0 列是標題，從第 1 列開始處理
            # A=0, B=1, C=2, F=5
            data_rows = df.iloc[1:].copy()
            
            # 邏輯過濾：B欄(1)不為空 且 F欄(5) >= 6
            # 先確保 F 欄是數字型態，無法轉換的變成 NaN
            data_rows[5] = pd.to_numeric(data_rows[5], errors='coerce')
            
            mask = (data_rows[1].notna()) & (data_rows[5] >= 6)
            filtered = data_rows[mask].copy()
            
            # 只取 A, B, C 欄位
            all_data.append(filtered[[0, 1, 2]])
            
        except Exception as e:
            st.error(f"檔案 {uploaded_file.name} 處理出錯: {e}")

    if all_data:
        # 合併所有符合條件的列
        final_df = pd.concat(all_data)
        final_df.columns = ['班級', '座號', '姓名']
        
        # 統計每個學生出現的次數 (根據 班級+座號+姓名 分組)
        summary = final_df.groupby(['班級', '座號', '姓名']).size().reset_index(name='達成區間數')
        
        # 排序：按班級、座號排序
        summary = summary.sort_values(by=['班級', '座號'])
        
        # 判定獎勵
        summary['獎勵狀態'] = summary['達成區間數'].apply(lambda x: "★ 精美禮物" if x >= 3 else "-")

        # 呈現結果
        st.subheader("📊 統計結果")
        
        # 設定表格樣式：達標者背景變色 (Streamlit 特色)
        def highlight_winners(s):
            return ['background-color: #ffcccc' if s['達成區間數'] >= 3 else '' for _ in s]

        st.dataframe(summary.style.apply(highlight_winners, axis=1), use_container_width=True)

        # 匯出 Excel 按鈕
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary.to_excel(writer, index=False, sheet_name='獎勵名單')
        
        st.download_button(
            label="📥 下載統計結果 Excel",
            data=output.getvalue(),
            file_name="閱讀獎勵統計結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("所選檔案中沒有符合條件 (B欄有資料且F欄>=6) 的數據。")