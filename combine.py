import streamlit as st
import pandas as pd
from string import ascii_uppercase

def col_index_to_letter(index):
    result = ""
    while index >= 0:
        result = ascii_uppercase[index % 26] + result
        index = index // 26 - 1
    return result

st.title("🔗 合并Excel文件")

# Create two columns
col1, col2 = st.columns(2)

with col1:
    st.header("📂 主文件")
    parent_file = st.file_uploader("上传主文件", type=["xlsx"], key="parent")
    if parent_file:
        parent_xl = pd.ExcelFile(parent_file)
        parent_sheet = st.selectbox("Select Parent Sheet", parent_xl.sheet_names, key="parent_sheet")
        parent_df = parent_xl.parse(parent_sheet)
        # st.dataframe(parent_df)
        st.dataframe(parent_df.rename(columns={col: col_index_to_letter(i) for i, col in enumerate(parent_df.columns)}))
        # parent_column = st.selectbox("Select Parent Column", [col_index_to_letter(i) for i in range(len(parent_df.columns))], key="parent_column")
        parent_key_col = st.selectbox("Select Parent Key Column", [col_index_to_letter(i) for i in range(len(parent_df.columns))], key="parent_key")
        parent_value_col = st.selectbox("Select Parent Value Column", [col_index_to_letter(i) for i in range(len(parent_df.columns))], key="parent_value")

with col2:
    st.header("📂 子文件")
    child_files = st.file_uploader("上传子文件", type=["xlsx"], accept_multiple_files=True, key="child")
    if child_files:
        selected_file = st.selectbox("Select One Child File to Preview", [f.name for f in child_files], key="child_file")
        selected_file_obj = next(f for f in child_files if f.name == selected_file)
        child_xl = pd.ExcelFile(selected_file_obj)
        child_sheet = st.selectbox("Select Child Sheet", child_xl.sheet_names, key="child_sheet")
        child_df = child_xl.parse(child_sheet)
        #  st.dataframe(child_df)
        st.dataframe(child_df.rename(columns={col: col_index_to_letter(i) for i, col in enumerate(child_df.columns)}))
        # child_column = st.selectbox("Select Child Column", [col_index_to_letter(i) for i in range(len(child_df.columns))], key="child_column")
        child_key_col = st.selectbox("Select Child Key Column", [col_index_to_letter(i) for i in range(len(child_df.columns))], key="child_key")
        child_value_col = st.selectbox("Select Child Value Column", [col_index_to_letter(i) for i in range(len(child_df.columns))], key="child_value")

if child_files and parent_file:
    st.subheader("📊 已提取的Key-Value对")
    results = []

    for file in child_files:
        xl = pd.ExcelFile(file)
        df = xl.parse(child_sheet)

        for idx, row in df.iterrows():
            val = row.iloc[int(ord(child_value_col) - ord("A"))]
            if pd.notna(val) and val != "":
                key = row.iloc[int(ord(child_key_col) - ord("A"))]
                results.append({"File": file.name, "Key": key, "Value": val})

    if results:
        st.dataframe(pd.DataFrame(results))

        # --- Begin new logic ---
        # Convert column letters to indices
        parent_key_idx = int(ord(parent_key_col) - ord("A"))
        parent_value_idx = int(ord(parent_value_col) - ord("A"))
        parent_data = parent_df.copy()

        # Build key-value lookup from parent
        parent_kv_map = {}
        for _, row in parent_data.iterrows():
            key = row.iloc[parent_key_idx]
            val = row.iloc[parent_value_idx]
            if pd.notna(key):
                parent_kv_map[key] = val

        # Track issues
        duplicates = {}
        new_keys = []

        for item in results:
            k = item["Key"]
            v = item["Value"]
            if k in parent_kv_map:
                if pd.notna(parent_kv_map[k]) and parent_kv_map[k] != v:
                    duplicates.setdefault(k, set()).add(parent_kv_map[k])
                    duplicates[k].add(v)
            else:
                new_keys.append((k, v))

        if duplicates:
            st.warning("⚠️ 多个值找到相同的Key:")
            for k, vals in duplicates.items():
                st.text(f"Key: {k} | Conflicting values: {', '.join(map(str, vals))}")

        if new_keys:
            st.warning("⚠️ 子文件中的某些Key不存在于主文件中，请手动检查:")
            for k, _ in new_keys:
                st.text(f"- {k}")
        # --- End new logic ---
        # Fill in missing values in parent_data from results
        filled_count = 0
        for item in results:
            k = item["Key"]
            v = item["Value"]
            for idx, row in parent_data.iterrows():
                if row.iloc[parent_key_idx] == k and pd.isna(row.iloc[parent_value_idx]):
                    parent_data.iat[idx, parent_value_idx] = v
                    filled_count += 1

        if filled_count:
            st.success(f"✅ 在主文件中填充了 {filled_count} 个值。")
            import io
            extracted_df = pd.DataFrame(results)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                extracted_df.to_excel(writer, index=False)
            st.download_button(
                label="📥 下载已提取的Key-Value对",
                data=output.getvalue(),
                file_name="extracted_key_value_pairs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info(f"💡 可在Excel中使用如下VLOOKUP公式填充主文件中的值：\n \n =VLOOKUP({parent_key_col}, [extracted_key_value_pairs.xlsx]Sheet1!$B:$C, 2, FALSE)")
    else:
        st.info("没有找到非空的Key-Value对。")