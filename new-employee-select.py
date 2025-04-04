# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# 页面基础设置
st.set_page_config(
    page_title="这个月有哪些员工新入职？",
    layout="centered",
    page_icon="📊"
)

# 固定说明文字
st.markdown("""**本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系**""")

# ====================== 配置参数 ======================
SPECIAL_IDS = {"31049588", "31268163"}  # 特殊员工系统号
EXCLUDE_DEPT = "证照支持部"              # 排除部门

# ====================== 功能函数 ======================
def validate_data(df):
    """数据格式校验"""
    required_columns = {'三级组织', '员工系统号', '姓名', '花名', '入职日期', '员工二级类别', '四级组织'}
    if not required_columns.issubset(df.columns):
        missing = required_columns - set(df.columns)
        return False, f"缺失必要字段：{', '.join(missing)}"
    try:
        pd.to_datetime(df['入职日期'])
    except:
        return False, "入职日期格式异常"
    return True, ""

def get_month_range(year, month):
    """获取月份首末日期"""
    if month == 12:
        return datetime(year, 12, 1), datetime(year, 12, 31)
    else:
        return (datetime(year, month, 1), 
                datetime(year, month+1, 1) - pd.Timedelta(days=1))

# ====================== 界面组件 ======================
uploaded_file = st.file_uploader(
    "📤 上传花名册数据（仅支持.xlsx格式）", 
    type=["xlsx"],
    help="请上传最新版本的员工花名册Excel文件"
)

# 创建并排布局
col_year, col_month = st.columns(2)
with col_year:
    selected_year = st.selectbox(
        "选择年份",
        options=range(2020, 2031),
        index=5,
        format_func=lambda x: f"{x}年"
    )
with col_month:
    selected_month = st.selectbox(
        "选择月份",
        options=range(1, 13),
        index=2,
        format_func=lambda x: f"{x}月"
    )

# ====================== 主处理流程 ======================
if st.button("🚀 开始分析", type="primary") and uploaded_file:
    # 数据加载与校验
    try:
        df = pd.read_excel(uploaded_file, sheet_name="花名册")
    except Exception as e:
        st.error(f"文件读取失败：{str(e)}")
        st.stop()
    
    # 执行数据校验
    is_valid, msg = validate_data(df)
    if not is_valid:
        st.error(f"数据校验失败：{msg}")
        st.stop()
    
    # 显示处理进度
    progress_bar = st.progress(0)
    status_msg = st.empty()
    
    # ===== 第一阶段处理 =====
    status_msg.markdown("**▶ 正在进行数据筛选...**")
    progress_bar.progress(30)
    
    # 日期处理
    df["入职日期"] = pd.to_datetime(df["入职日期"])
    start_date, end_date = get_month_range(selected_year, selected_month)
    
    # 构建筛选条件
    mask = (
        df["入职日期"].between(start_date, end_date) &
        (df["员工二级类别"] == "正式员工") &
        (df["四级组织"] != EXCLUDE_DEPT) &
        (~df["员工系统号"].astype(str).isin(SPECIAL_IDS))
    )
    
    # 执行筛选
    filtered_df = df[mask].copy()
    result_df = filtered_df[["三级组织", "员工系统号", "姓名", "花名", "入职日期", "员工二级类别"]]
    result_df = result_df.sort_values(by=["三级组织", "入职日期"], ascending=[False, True])
    
    # ===== 第二阶段处理 =====
    status_msg.markdown("**▶ 正在生成汇总报告...**")
    progress_bar.progress(70)
    
    # 生成拼接字段
    result_df["姓名+花名"] = result_df.apply(
        lambda x: f"{x['姓名']}（{x['花名']}）" if pd.notnull(x['花名']) else x['姓名'],
        axis=1
    )
    
    # 执行分组聚合
    grouped_df = result_df.groupby("三级组织")["姓名+花名"].agg(
        lambda x: "、".join(x)
    ).reset_index()
    
    # ===== 生成输出文件 =====
    status_msg.markdown("**▶ 正在准备下载文件...**")
    progress_bar.progress(90)
    
    # 生成内存文件对象
    output1 = BytesIO()
    result_df.to_excel(output1, index=False)
    output1.seek(0)
    
    output2 = BytesIO()
    grouped_df.to_excel(output2, index=False)
    output2.seek(0)
    
    # ===== 显示最终结果 =====
    progress_bar.progress(100)
    status_msg.empty()
    progress_bar.empty()
    
    st.success("✅ 分析完成！")
    
    # 显示统计信息
    st.metric("符合条件员工总数", len(result_df))
    
    # 显示特殊提醒
    st.markdown("""
    ​**🔔 请人工检查以下情况：​**
    - 特殊原因外包人员
    - 活水人员（跨组织调动）
    """)
    
    # 显示被排除的特殊人员
    excluded = df[~mask & df["员工系统号"].astype(str).isin(SPECIAL_IDS)]
    if not excluded.empty:
        st.warning(f"已排除特殊人员：{', '.join(excluded['姓名'].tolist())}")
    
    # 下载按钮布局
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="⬇️ 下载保留人员明细",
            data=output1.getvalue(),
            file_name=f"保留人员明细_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_dl2:
        st.download_button(
            label="⬇️ 下载拼接结果",
            data=output2.getvalue(),
            file_name=f"人员信息拼接_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif not uploaded_file and st.button("🚀 开始分析"):
    st.warning("⚠️ 请先上传花名册数据文件")
