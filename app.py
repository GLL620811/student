import streamlit as st
import pandas as pd

# 读取 Excel 文件
file_path = "D:/昆明城市学院/昆明城市学院计科研究院/2025.06.11-KMCC-0929.xlsx"  # 修改为你的文件路径
df = pd.read_excel(file_path)

st.title("🎓 学生成绩筛选系统")

st.sidebar.header("筛选条件")

# ================= 考试等级按钮 =================
exam_level = st.sidebar.radio("考试等级选择", ["不选", "CET4", "CET6"])
filtered_df = df.copy()
if exam_level != "不选":
    filtered_df = filtered_df[filtered_df["考试等级"] == exam_level]

# ================= 通用区间筛选函数 =================
def apply_range_filter(df, column, label, default_range):
    """用户可选是否添加区间过滤"""
    use_filter = st.sidebar.checkbox(f"是否筛选 {label}", value=False)
    if not use_filter:
        return df

    ranges_text = st.sidebar.text_area(
        f"{label} 区间设置 (例如: {default_range})", default_range
    )
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

# ================= 应用数值筛选 =================
range_defaults = {
    "综合成绩": "0-710",
    "听力": "0-249",
    "阅读": "0-249",
    "写作": "0-212",
    "翻译": "0-212"
}

for col, default_range in range_defaults.items():
    if col in df.columns:  # 防止列不存在时报错
        filtered_df = apply_range_filter(filtered_df, col, col, default_range)

# ================= CET4 / CET6 考试次数 =================
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

# ================= 显示结果 =================
st.subheader("筛选结果")
st.dataframe(filtered_df)

# ================= 导出 Excel =================
if st.button("导出结果到Excel"):
    save_path = "筛选结果.xlsx"
    filtered_df.to_excel(save_path, index=False)
    st.success(f"✅ 已导出到 {save_path}")
