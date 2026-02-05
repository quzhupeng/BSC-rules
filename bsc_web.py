#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
平衡计分卡KPI数据处理 Web 应用
基于 Streamlit 的用户界面
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import openpyxl

# 导入核心处理类
from bsc_core import BSCProcessor

# 页面配置
st.set_page_config(
    page_title="BSC计分规则处理器",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
    }
    /* 高亮人工校验行的样式 */
    .stDataFrame[data-testid="stDataFrame"] div[data-testid="stDataFrameContainer"] {
        overflow-x: auto;
    }
</style>
""", unsafe_allow_html=True)

# 应用标题
st.markdown('<h1 class="main-header">📊 平衡计分卡 KPI 数据处理器</h1>', unsafe_allow_html=True)

# 侧边栏说明
with st.sidebar:
    st.image("https://img.icons8.com/color/96/spreadsheet.png", width=80)
    st.title("功能说明")
    st.info("""
    本工具用于将非结构化的KPI考核指标数据转化为标准化的平衡计分卡格式。

    **支持的功能：**
    - 自动识别目标值列和计分规则列
    - 数据清洗（百分比格式统一）
    - 底线值智能推导
    - 指标方向判定
    - 规范化计分规则生成

    **使用方法：**
    1. 上传包含KPI数据的Excel文件
    2. 等待自动处理完成
    3. 预览处理结果
    4. 下载处理后的文件
    """)

    st.markdown("---")
    st.markdown("**支持的计分规则类型：**")
    st.markdown("""
    - 📉 每低X%扣Y分
    - 🔢 每少X个扣Y分
    - 📊 实际/目标×100
    - ⚠️ 显式阈值声明
    """)

# 初始化session state
if 'processed_df' not in st.session_state:
    st.session_state.processed_df = None
if 'processor' not in st.session_state:
    st.session_state.processor = None
if 'stats' not in st.session_state:
    st.session_state.stats = None
if 'logs' not in st.session_state:
    st.session_state.logs = []

# 文件上传区域
st.markdown("### 📁 文件上传")
uploaded_file = st.file_uploader(
    "请上传Excel文件 (.xlsx)",
    type=['xlsx', 'xls'],
    label_visibility="collapsed",
    help="上传包含目标值和计分规则列的Excel文件"
)

# 处理按钮
if uploaded_file is not None:
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"**已选择文件：** `{uploaded_file.name}`")
    with col2:
        if st.button("🚀 开始处理", type="primary", use_container_width=True):
            with st.spinner("正在处理数据，请稍候..."):
                try:
                    # 读取文件到BytesIO
                    file_bytes = BytesIO(uploaded_file.getvalue())
                    file_bytes.name = uploaded_file.name

                    # 创建处理器
                    processor = BSCProcessor(file_bytes)

                    # 进度条
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    def progress_callback(progress):
                        progress_bar.progress(progress)
                        if progress < 40:
                            status_text.text("正在读取文件...")
                        elif progress < 50:
                            status_text.text("正在识别列...")
                        elif progress < 95:
                            status_text.text("正在处理数据...")
                        else:
                            status_text.text("处理完成！")

                    # 执行处理
                    result_df = processor.process(progress_callback)

                    # 保存到session state
                    st.session_state.processed_df = result_df
                    st.session_state.processor = processor
                    st.session_state.stats = processor.get_stats()
                    st.session_state.logs = processor.get_logs()

                    progress_bar.progress(100)
                    status_text.text("✅ 处理完成！")

                    st.success("处理成功！请查看下方结果。")

                except Exception as e:
                    st.error(f"处理失败：{str(e)}")
                    st.exception(e)

# 显示处理结果
if st.session_state.processed_df is not None:
    st.markdown("---")
    st.markdown("### 📈 处理结果")

    # 统计信息
    stats = st.session_state.stats
    if stats:
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("总行数", stats['total'])
        with col2:
            st.metric("✅ 成功解析", stats['success'], delta_color="normal")
        with col3:
            st.metric("⚠️ 需人工校验", stats['manual_check'], delta_color="inverse")
        with col4:
            if stats['error'] > 0:
                st.metric("❌ 错误", stats['error'])
            else:
                st.metric("错误", stats['error'])

    # 处理日志
    if st.session_state.logs:
        with st.expander("📋 查看处理日志"):
            for log in st.session_state.logs:
                st.text(log)

    # 数据预览
    st.markdown("### 📋 数据预览")

    df = st.session_state.processed_df

    # 获取要显示的列
    display_columns = [col for col in df.columns if not col.startswith('_') and col not in ['目标值_归一化', '底线值_归一化', '是否百分比']]

    # 高亮人工校验行的样式函数
    def highlight_manual_check(row):
        if row.get('解析状态') == '人工校验':
            return ['background-color: #fff3cd'] * len(row)
        elif row.get('解析状态', '').startswith('ERROR'):
            return ['background-color: #f8d7da'] * len(row)
        return [''] * len(row)

    # 应用样式
    styled_df = df[display_columns].style.apply(highlight_manual_check, axis=1)

    # 显示数据
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=400
    )

    # 颜色说明
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<span style="background-color: #fff3cd; padding: 4px 12px; border-radius: 4px;">⚠️ 黄色背景 = 需人工校验</span>', unsafe_allow_html=True)
    with col2:
        st.markdown('<span style="background-color: #f8d7da; padding: 4px 12px; border-radius: 4px;">❌ 红色背景 = 解析错误</span>', unsafe_allow_html=True)

    # 下载按钮
    st.markdown("---")
    st.markdown("### 💾 下载结果")

    col1, col2, col3 = st.columns([1, 1, 2])

    with col1:
        # 生成Excel文件
        excel_data = st.session_state.processor.save_to_bytesio()

        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"BSC处理结果_{timestamp}.xlsx"

        st.download_button(
            label="📥 下载 Excel 文件",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        # 同时提供CSV下载选项
        csv_data = df[display_columns].to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📄 下载 CSV 文件",
            data=csv_data,
            file_name=f"BSC处理结果_{timestamp}.csv",
            mime="text/csv",
            use_container_width=True
        )

# 底部信息
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #6c757d; font-size: 0.9rem;">
        平衡计分卡KPI数据处理器 v1.0 | 基于 Streamlit 构建
    </div>
    """,
    unsafe_allow_html=True
)

# 空状态提示
if uploaded_file is None and st.session_state.processed_df is None:
    st.markdown("---")
    st.markdown("""
    ### 👋 欢迎使用平衡计分卡KPI数据处理器

    请在上方上传您的Excel文件开始处理。

    **文件要求：**
    - Excel文件需包含 **目标值列**（列名包含"目标值"关键字）
    - Excel文件需包含 **计分规则列**（列名包含"计分规则"关键字）

    如有问题，请查看左侧功能说明。
    """)
