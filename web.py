
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Merge Tool · Licht", layout="wide")

st.title("Excel Merge Tool")
st.markdown("Merge columns from Table 1 into Table.")

def read_excel_smart(uploaded_file, header_row, sheet_name=0):
    try:
        df = pd.read_excel(uploaded_file, header=header_row, sheet_name=sheet_name)
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique():
            dup_idx = cols[cols == dup].index.tolist()
            cols[dup_idx] = [dup + '_' + str(i) if i != 0 else dup for i in range(len(dup_idx))]
        df.columns = cols
        df = df.dropna(how='all').reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"读取文件失败: {e}")
        return None

def is_cell_empty(val):
    if pd.isna(val):
        return True
    v_str = str(val).strip()
    return v_str == '' or v_str.lower() in ('nan', 'none', 'null', 'nat')

def make_match_key(row, cols):
    vals = []
    for c in cols:
        v = str(row[c]).strip() if pd.notna(row[c]) else ''
        vals.append(v)
    if any(v in ('', 'nan', 'none', 'null') for v in vals):
        return None
    return '|||'.join(vals)

def reset_all():
    for k in ['df1_raw', 'df1_clean', 'df2_raw', 'df2_clean', 'match_cols_1', 'merge_cols_1',
              'match_cols_2', 'merge_mapping', 'result_df', 'header_row_1', 'header_row_2',
              'last_file1', 'last_file2']:
        if k in st.session_state:
            del st.session_state[k]

for key in ['df1_raw', 'df1_clean', 'df2_raw', 'df2_clean', 'match_cols_1', 'merge_cols_1',
            'match_cols_2', 'merge_mapping', 'result_df', 'header_row_1', 'header_row_2',
            'last_file1', 'last_file2']:
    if key not in st.session_state:
        st.session_state[key] = None

st.header("📁 1-上传并配置表1（源表）")

uploaded_file_1 = st.file_uploader("上传表1（Excel格式）", type=["xlsx", "xls"], key="file1")

if uploaded_file_1 is not None:
    if st.session_state.get('last_file1') != uploaded_file_1.name:
        reset_all()
        st.session_state['last_file1'] = uploaded_file_1.name
        st.rerun()

    col1, col2 = st.columns([1, 3])
    with col1:
        header_row_1 = st.number_input("表1表头所在行（从0开始）", min_value=0, max_value=20,
                                       value=st.session_state.get('header_row_1', 0), key="header1")
        st.session_state['header_row_1'] = header_row_1

    df1_raw = read_excel_smart(uploaded_file_1, int(header_row_1))

    if df1_raw is not None:
        st.session_state.df1_raw = df1_raw
        with col2:
            st.success(f"表1读取成功：{len(df1_raw)} 行 × {len(df1_raw.columns)} 列")

        with st.expander("🔍 表1数据预览", expanded=False):
            st.dataframe(df1_raw.head(20), use_container_width=True)

        st.markdown("---")
        st.subheader("选择表1的匹配列与融合列")

        col_a, col_b = st.columns(2)
        with col_a:
            match_cols_1 = st.multiselect(
                "🔑 选择表1的匹配列",
                options=df1_raw.columns.tolist(),
                default=st.session_state.match_cols_1 or [],
                key="match_cols_1_select",
                help="用于在表2中查找对应行的列，可多选"
            )
            st.session_state.match_cols_1 = match_cols_1

        with col_b:
            available = [c for c in df1_raw.columns.tolist() if c not in match_cols_1]
            merge_cols_1 = st.multiselect(
                "📋 选择表1的融合列",
                options=available,
                default=st.session_state.merge_cols_1 or [],
                key="merge_cols_1_select",
                help="要合并到表2的列，可多选"
            )
            st.session_state.merge_cols_1 = merge_cols_1

        if match_cols_1:
            mask = df1_raw[match_cols_1].notna().all(axis=1)
            for col in match_cols_1:
                mask = mask & (~df1_raw[col].astype(str).str.strip().str.lower().isin(['nan', 'none', 'null', '']))
            df1_clean = df1_raw[mask].reset_index(drop=True)
            st.session_state.df1_clean = df1_clean
            st.info(f"✅ 表1有效数据：{len(df1_clean)} 行")
        else:
            st.info("👆 请先选择至少一个匹配列")

if (st.session_state.df1_clean is not None and 
    st.session_state.match_cols_1 and 
    st.session_state.merge_cols_1):

    st.markdown("---")
    st.header("📁 2-上传并配置表2（目标表）")

    uploaded_file_2 = st.file_uploader("上传表2（Excel格式）", type=["xlsx", "xls"], key="file2")

    if uploaded_file_2 is not None:
        if st.session_state.get('last_file2') != uploaded_file_2.name:
            for k in ['df2_raw', 'df2_clean', 'match_cols_2', 'merge_mapping', 'result_df', 'header_row_2', 'last_file2']:
                st.session_state[k] = None
            st.session_state['last_file2'] = uploaded_file_2.name
            st.rerun()

        col1, col2 = st.columns([1, 3])
        with col1:
            header_row_2 = st.number_input("表2表头所在行（从0开始）", min_value=0, max_value=20,
                                           value=st.session_state.get('header_row_2', 0), key="header2")
            st.session_state['header_row_2'] = header_row_2

        df2_raw = read_excel_smart(uploaded_file_2, int(header_row_2))

        if df2_raw is not None:
            st.session_state.df2_raw = df2_raw
            with col2:
                st.success(f"表2读取成功：{len(df2_raw)} 行 × {len(df2_raw.columns)} 列")

            with st.expander("🔍 表2数据预览", expanded=False):
                st.dataframe(df2_raw.head(20), use_container_width=True)

            st.markdown("---")
            st.subheader("配置表2的匹配列")

            match_cols_2 = []
            match_valid = True

            cols_match = st.columns(len(st.session_state.match_cols_1))
            for i, col1_name in enumerate(st.session_state.match_cols_1):
                with cols_match[i]:
                    selected = st.selectbox(
                        f"对应表1「{col1_name}」",
                        options=["-- 请选择 --"] + df2_raw.columns.tolist(),
                        key=f"match_col_2_{i}"
                    )
                    if selected != "-- 请选择 --":
                        match_cols_2.append(selected)
                    else:
                        match_valid = False

            if match_valid and len(match_cols_2) == len(st.session_state.match_cols_1):
                mask = df2_raw[match_cols_2].notna().all(axis=1)
                for col in match_cols_2:
                    mask = mask & (~df2_raw[col].astype(str).str.strip().str.lower().isin(['nan', 'none', 'null', '']))
                df2_clean = df2_raw[mask].reset_index(drop=True)
                st.session_state.df2_clean = df2_clean
                st.success(f"✅ 匹配列配置完成！表2有效数据：{len(df2_clean)} 行")
                st.session_state.match_cols_2 = match_cols_2

                with st.expander("🔍 表2清理后预览", expanded=False):
                    st.dataframe(df2_clean.head(20), use_container_width=True)

                df1_names = set()
                for _, row in st.session_state.df1_clean.iterrows():
                    key = make_match_key(row, st.session_state.match_cols_1)
                    if key:
                        df1_names.add(key)
                df2_names = set()
                for _, row in df2_clean.iterrows():
                    key = make_match_key(row, match_cols_2)
                    if key:
                        df2_names.add(key)

                only_in_1 = df1_names - df2_names
                only_in_2 = df2_names - df1_names
                both = df1_names & df2_names

                col_stat1, col_stat2, col_stat3 = st.columns(3)
                with col_stat1:
                    st.metric("两表共有", len(both))
                with col_stat2:
                    st.metric("仅在表1（将追加）", len(only_in_1))
                with col_stat3:
                    st.metric("仅在表2", len(only_in_2))

                if only_in_1:
                    with st.expander(f"📋 将追加到表2的行（{len(only_in_1)}个）", expanded=False):
                        # 显示这些行的匹配键
                        st.write("以下匹配键在表1中存在，但表2中不存在，将被追加：")
                        for k in sorted(only_in_1):
                            st.write(f"- {k.replace('|||', ' | ')}")

                # ---- 融合列配置 ----
                st.markdown("---")
                st.subheader("📋 表2融合列配置（与表1融合列一一对应）")

                merge_mapping = {}
                merge_valid = True

                for i, col1_name in enumerate(st.session_state.merge_cols_1):
                    with st.container(border=True):
                        col_left, col_mid, col_right = st.columns([1.2, 1.5, 1.5])
                        with col_left:
                            st.markdown(f"**表1：「{col1_name}」**")

                        with col_mid:
                            options = ["➕ 新建列"] + df2_clean.columns.tolist()
                            default_idx = 0
                            # 如果表2有同名列，默认选中
                            if col1_name in df2_clean.columns:
                                default_idx = options.index(col1_name)

                            selected = st.selectbox(
                                f"表2目标列",
                                options=options,
                                index=default_idx,
                                key=f"merge_col_2_{i}"
                            )

                        with col_right:
                            if selected == "➕ 新建列":
                                suggested = col1_name
                                if suggested in df2_clean.columns:
                                    suggested = f"{col1_name}_新"
                                new_name = st.text_input(
                                    f"新列名称",
                                    value=suggested,
                                    key=f"merge_new_name_{i}"
                                )
                                if new_name.strip():
                                    merge_mapping[col1_name] = new_name.strip()
                                    st.caption(f"✅ 新建列「{new_name.strip()}」")
                                else:
                                    merge_valid = False
                                    st.caption("⚠️ 请输入列名")
                            else:
                                merge_mapping[col1_name] = selected
                                # 统计该列空值情况
                                empty_count = sum(1 for v in df2_clean[selected] if is_cell_empty(v))
                                st.caption(f"该列空值 {empty_count}/{len(df2_clean)} 行")

                st.session_state.merge_mapping = merge_mapping

                # 追加行时其他列的处理选项
                st.markdown("---")
                st.subheader("⚙️ 追加行配置")
                fill_match_cols = st.checkbox("追加行时，匹配列也填入表1的值", value=True, key="fill_match")

                if merge_valid and len(merge_mapping) == len(st.session_state.merge_cols_1):
                    mapping_df = pd.DataFrame({
                        "表1融合列": list(merge_mapping.keys()),
                        "表2目标列": list(merge_mapping.values()),
                        "操作": ["新建列" if col not in df2_clean.columns else "填充/追加" for col in merge_mapping.values()]
                    })
                    st.dataframe(mapping_df, use_container_width=True, hide_index=True)

                    # ===================== 步骤3：执行合并 =====================
                    st.markdown("---")
                    st.header("🚀 步骤三：执行合并")

                    if st.button("▶️ 开始合并", type="primary", use_container_width=True):
                        with st.spinner("正在合并数据..."):
                            df1 = st.session_state.df1_clean.copy()
                            df2 = st.session_state.df2_clean.copy()
                            match_cols_1 = st.session_state.match_cols_1
                            match_cols_2 = st.session_state.match_cols_2
                            merge_mapping = st.session_state.merge_mapping

                            # 转为字符串
                            for col in match_cols_1:
                                df1[col] = df1[col].astype(str).str.strip()
                            for col in match_cols_2:
                                df2[col] = df2[col].astype(str).str.strip()

                            # 创建匹配键
                            df1['_match_key'] = df1.apply(lambda row: make_match_key(row, match_cols_1), axis=1)
                            df2['_match_key'] = df2.apply(lambda row: make_match_key(row, match_cols_2), axis=1)

                            # 确保所有目标列存在
                            for src_col, tgt_col in merge_mapping.items():
                                if tgt_col not in df2.columns:
                                    df2[tgt_col] = None

                            # 建立表1查找字典：match_key -> 完整行数据
                            lookup = {}
                            for _, row in df1.iterrows():
                                key = row['_match_key']
                                if key is None:
                                    continue
                                lookup[key] = row.to_dict()

                            # 统计
                            filled_count = 0      # 表2已有行，空白单元格被填充
                            skipped_count = 0     # 表2已有行，已有内容被跳过
                            appended_count = 0    # 表1有但表2没有，追加新行

                            # 第一步：处理表2已有行
                            for idx, row in df2.iterrows():
                                key = row['_match_key']
                                if key is None or key not in lookup:
                                    continue

                                for src_col, tgt_col in merge_mapping.items():
                                    val = lookup[key].get(src_col)
                                    current_val = row[tgt_col]

                                    if is_cell_empty(current_val):
                                        if not is_cell_empty(val):
                                            df2.at[idx, tgt_col] = val
                                            filled_count += 1
                                    else:
                                        skipped_count += 1

                            # 第二步：追加表1有但表2没有的行
                            df2_keys = set(df2['_match_key'].dropna().tolist())

                            new_rows = []
                            for key, row_data in lookup.items():
                                if key not in df2_keys:
                                    new_row = {}
                                    # 表2原有列全部设为空
                                    for col in df2.columns:
                                        if col != '_match_key':
                                            new_row[col] = None

                                    # 填入匹配列（如果配置了）
                                    if fill_match_cols:
                                        for c1, c2 in zip(match_cols_1, match_cols_2):
                                            new_row[c2] = row_data.get(c1)

                                    # 填入融合列
                                    for src_col, tgt_col in merge_mapping.items():
                                        val = row_data.get(src_col)
                                        if not is_cell_empty(val):
                                            new_row[tgt_col] = val

                                    new_rows.append(new_row)
                                    appended_count += 1

                            if new_rows:
                                new_rows_df = pd.DataFrame(new_rows)
                                # 确保列顺序一致
                                for col in df2.columns:
                                    if col not in new_rows_df.columns and col != '_match_key':
                                        new_rows_df[col] = None
                                new_rows_df = new_rows_df[[c for c in df2.columns if c != '_match_key']]
                                df2 = pd.concat([df2, new_rows_df], ignore_index=True)

                            df2 = df2.drop(columns=['_match_key'], errors='ignore')
                            st.session_state.result_df = df2

                            # 显示结果
                            st.success(f"✅ 合并完成！")

                            col_s1, col_s2, col_s3, col_s4 = st.columns(4)
                            with col_s1:
                                st.metric("原表2行数", len(df2) - appended_count)
                            with col_s2:
                                st.metric("追加行数", appended_count)
                            with col_s3:
                                st.metric("填充单元格", filled_count)
                            with col_s4:
                                st.metric("跳过已有内容", skipped_count)

                            with st.expander("🔍 合并结果预览", expanded=True):
                                st.dataframe(df2, use_container_width=True)

                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df2.to_excel(writer, index=False, sheet_name='合并结果')

                            st.download_button(
                                label="📥 下载合并后的Excel文件",
                                data=output.getvalue(),
                                file_name="合并结果.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                else:
                    st.warning("⚠️ 请完成所有融合列的配置")
            else:
                st.warning("⚠️ 请为每个表1匹配列选择对应的表2列")
else:
    if st.session_state.df1_raw is None:
        st.info("👆 请先上传表1")
    elif not st.session_state.match_cols_1:
        st.info("👆 请在表1中选择至少一个匹配列")
    elif not st.session_state.merge_cols_1:
        st.info("👆 请在表1中选择至少一个融合列")

st.markdown("---")
st.caption("💡 使用说明：")
st.caption("1. 匹配列用于在表1和表2之间建立对应关系。")
st.caption("2. 融合列是表1中要合并到表2的数据列。")
st.caption("3. 表2中已有行：融合列为空则填充，已有内容不覆盖。")
st.caption("4. 表1有但表2没有的行：会自动追加到表2末尾，融合列填入数据，其他列留空。")
st.caption("5. 如果表2没有对应融合列，可选择「➕ 新建列」并自定义列名。")