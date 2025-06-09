import streamlit as st
import pandas as pd
from string import ascii_uppercase
import time

@st.cache_resource
def load_excel_file_from_bytes(file_bytes):
    return pd.ExcelFile(file_bytes)

@st.cache_data
def parse_excel_sheet_from_bytes(file_bytes, sheet_name):
    xl = load_excel_file_from_bytes(file_bytes)
    return xl.parse(sheet_name)

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
        start = time.time()
        st.write("📥 正在读取文件字节...")
        parent_bytes = parent_file.getvalue()
        st.write(f"✅ 文件读取完成，用时 {time.time() - start:.2f} 秒")

        start = time.time()
        st.write("🔍 正在解析 ExcelFile...")
        parent_xl = load_excel_file_from_bytes(parent_bytes)
        st.write(f"✅ ExcelFile 解析完成，用时 {time.time() - start:.2f} 秒")

        parent_sheet = st.selectbox("选择主文件的Sheet", parent_xl.sheet_names, key="parent_sheet")
        parent_df = parse_excel_sheet_from_bytes(parent_bytes, parent_sheet)
        # st.dataframe(parent_df)
        st.dataframe(parent_df.rename(columns={col: col_index_to_letter(i) for i, col in enumerate(parent_df.columns)}))
        # parent_column = st.selectbox("Select Parent Column", [col_index_to_letter(i) for i in range(len(parent_df.columns))], key="parent_column")
        parent_key_col = st.selectbox("选择主文件的Key列", [col_index_to_letter(i) for i in range(len(parent_df.columns))], key="parent_key")
        parent_value_col = st.selectbox("选择主文件的Value列", [col_index_to_letter(i) for i in range(len(parent_df.columns))], key="parent_value")

with col2:
    st.header("📂 子文件")
    child_files = st.file_uploader("上传子文件", type=["xlsx"], accept_multiple_files=True, key="child")
    if child_files:
        selected_file = st.selectbox("选择一个子文件", [f.name for f in child_files], key="child_file")
        selected_file_obj = next(f for f in child_files if f.name == selected_file)

        start = time.time()
        st.write("📥 正在读取子文件字节...")
        child_bytes = selected_file_obj.getvalue()
        st.write(f"✅ 子文件读取完成，用时 {time.time() - start:.2f} 秒")

        start = time.time()
        st.write("🔍 正在解析子文件 ExcelFile...")
        child_xl = load_excel_file_from_bytes(child_bytes)
        st.write(f"✅ 子文件解析完成，用时 {time.time() - start:.2f} 秒")

        default_child_sheet = parent_sheet if parent_file and parent_sheet in child_xl.sheet_names else child_xl.sheet_names[0]
        child_sheet = st.selectbox(
            "选择子文件的Sheet",
            child_xl.sheet_names,
            index=child_xl.sheet_names.index(default_child_sheet),
            key="child_sheet"
        )
        child_df = parse_excel_sheet_from_bytes(child_bytes, child_sheet)
        #  st.dataframe(child_df)
        st.dataframe(child_df.rename(columns={col: col_index_to_letter(i) for i, col in enumerate(child_df.columns)}))
        # child_column = st.selectbox("Select Child Column", [col_index_to_letter(i) for i in range(len(child_df.columns))], key="child_column")
        child_col_options = [col_index_to_letter(i) for i in range(len(child_df.columns))]
        default_child_key = parent_key_col if 'parent_key_col' in locals() and parent_key_col in child_col_options else child_col_options[0]
        default_child_value = parent_value_col if 'parent_value_col' in locals() and parent_value_col in child_col_options else child_col_options[0]
        child_key_col = st.selectbox(
            "选择子文件的Key列",
            child_col_options,
            index=child_col_options.index(default_child_key),
            key="child_key"
        )
        child_value_col = st.selectbox(
            "选择子文件的Value列",
            child_col_options,
            index=child_col_options.index(default_child_value),
            key="child_value"
        )

if child_files and parent_file:
    st.subheader("📊 已提取的Key-Value对")
    results = []

    for file in child_files:
        file_bytes = file.getvalue()
        df = parse_excel_sheet_from_bytes(file_bytes, child_sheet)

        for idx, row in df.iterrows():
            val = row.iloc[int(ord(child_value_col) - ord("A"))]
            if pd.notna(val) and val != "":
                key = row.iloc[int(ord(child_key_col) - ord("A"))]
                results.append({"文件": file.name, "Key": key, "Value": val})

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
                st.text(f"Key: {k} | 冲突的 Value: {', '.join(map(str, vals))}")

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
            st.info(f"💡 可在Excel中使用如下VLOOKUP公式填充主文件中的值：\n \n =IFERROR(VLOOKUP({parent_key_col}, [extracted_key_value_pairs.xlsx]Sheet1!B:C, 2, FALSE), '')")
    else:
        st.info("没有找到非空的Key-Value对。")