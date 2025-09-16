import io
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder

# ================== 缓存读取 Excel ==================
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    return df

file_path = "student.xlsx"
df = load_data(file_path)

st.title("🎓 学生成绩筛选系统")
st.sidebar.header("筛选条件")

# ================== 考试等级选择 ==================
exam_level = st.sidebar.radio("考试等级选择", ["不选", "CET4", "CET6"])
filtered_df = df.copy()
if exam_level != "不选":
    filtered_df = filtered_df[filtered_df["考试等级"] == exam_level]

# ================== CET4 / CET6 分数区间 ==================
def apply_score_range(df, exam_type):
    if "综合成绩" not in df.columns:
        return df
    score_range = st.sidebar.radio(
        f"{exam_type} 综合成绩区间选择",
        ["不选", "425分以上", "400-425分", "400分以下"]
    )
    if score_range != "不选":
        if score_range == "425分及以上":
            df = df[df["综合成绩"] >= 425]
        elif score_range == "400-425分":
            df = df[(df["综合成绩"] >= 400) & (df["综合成绩"] < 425)]
        elif score_range == "400分以下":
            df = df[df["综合成绩"] < 400]
    return df

if exam_level == "CET4":
    filtered_df = apply_score_range(filtered_df, "CET4")
elif exam_level == "CET6":
    filtered_df = apply_score_range(filtered_df, "CET6")

# ================== 通用区间筛选 ==================
def apply_range_filter(df, column, label, default_range):
    use_filter = st.sidebar.checkbox(f"是否筛选 {label}", value=False)
    if not use_filter:
        return df
    ranges_text = st.sidebar.text_area(f"{label} 区间设置 (例如: {default_range})", default_range)
    if not ranges_text.strip():
        return df

    ranges = []
    for r in ranges_text.split(","):
        r = r.strip()
        if "-" in r:
            try:
                start, end = map(int, r.split("-"))
                ranges.append((start, end))
            except:
                st.sidebar.error(f"区间格式错误: {r}")

    if not ranges:
        return df

    mask = False
    for (start, end) in ranges:
        mask = mask | ((df[column] >= start) & (df[column] <= end))
    return df[mask]

range_defaults = {
    "听力": "0-249",
    "阅读": "0-249",
    "写作": "0-212",
    "翻译": "0-212"
}

for col, default_range in range_defaults.items():
    if col in df.columns:
        filtered_df = apply_range_filter(filtered_df, col, col, default_range)

# ================== CET4 / CET6 考试次数多选 ==================
def apply_multi_select(df, column, label):
    if column not in df.columns:
        return df
    unique_vals = sorted(df[column].dropna().unique().tolist())
    selected_vals = st.sidebar.multiselect(f"{label} (可多选)", unique_vals)
    if selected_vals:
        return df[df[column].isin(selected_vals)]
    return df

filtered_df = apply_multi_select(filtered_df, "CET4考试次数", "CET4考试次数")
filtered_df = apply_multi_select(filtered_df, "CET6考试次数", "CET6考试次数")

# ================== 显示结果（AgGrid 分页） ==================
st.subheader("筛选结果")
st.write(f"总共筛选到 {len(filtered_df)} 条记录")

gb = GridOptionsBuilder.from_dataframe(filtered_df)
gb.configure_pagination(paginationAutoPageSize=True)  # 自动分页
gb.configure_side_bar()  # 可展开侧边栏筛选
gridOptions = gb.build()

AgGrid(
    filtered_df,
    gridOptions=gridOptions,
    enable_enterprise_modules=False,
    allow_unsafe_jscode=True,
    fit_columns_on_grid_load=True
)

# ================== 导出 Excel ==================
if st.button("导出结果到Excel"):
    # 创建内存中的 BytesIO 缓冲区
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="筛选结果")
    excel_data = output.getvalue()

    # 下载按钮
    st.download_button(
        label="📥 点击下载 Excel 文件",
        data=excel_data,
        file_name="筛选结果.xlsx",  # 默认文件名，用户可自己改保存位置
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )
