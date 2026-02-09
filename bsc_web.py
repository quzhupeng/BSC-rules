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
from bsc_core import BSCProcessor, BSCMultiSheetProcessor, BSCBatchProcessor

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
    - **半年度数据同步处理**（自动识别半年度列）
    - **多Sheet同步处理**

    **使用方法：**
    1. 选择处理模式（单Sheet/多Sheet/批量文件）
    2. 上传Excel文件（批量模式支持多个文件）
    3. 等待自动处理完成
    4. 预览处理结果
    5. 下载处理后的文件
    """)

    st.markdown("---")
    st.markdown("**支持的计分规则类型：**")
    st.markdown("""
    - 📉 每低X%扣Y分
    - 🔢 每少X个扣Y分
    - 📊 实际/目标×100
    - ⚠️ 显式阈值声明
    - 📑 多级计分规则（XX得60分）
    """)

    st.markdown("---")
    st.markdown("**处理模式说明：**")
    st.markdown("""
    - **单Sheet处理**：只处理第一个有数据的Sheet
    - **多Sheet处理**：自动检测并处理所有包含KPI数据的Sheet，每个Sheet输出为结果文件中的一个Sheet
    - **批量文件处理**：一次上传多个Excel文件，自动处理所有Sheet，结果合并到一个Excel输出
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
if 'multi_sheet_processor' not in st.session_state:
    st.session_state.multi_sheet_processor = None
if 'multi_sheet_stats' not in st.session_state:
    st.session_state.multi_sheet_stats = None
if 'is_multi_sheet' not in st.session_state:
    st.session_state.is_multi_sheet = False
if 'batch_processor' not in st.session_state:
    st.session_state.batch_processor = None
if 'batch_stats' not in st.session_state:
    st.session_state.batch_stats = None
if 'is_batch' not in st.session_state:
    st.session_state.is_batch = False

# 文件上传区域
st.markdown("### 📁 文件上传")

# 处理模式选择（放在文件上传之前，因为 accept_multiple_files 在渲染时确定）
with st.columns([1, 1])[0]:
    processing_mode = st.radio(
        "处理模式",
        ["单Sheet处理", "多Sheet处理", "批量文件处理"],
        horizontal=True,
        help="单Sheet: 只处理第一个有数据的Sheet | 多Sheet: 处理所有包含KPI数据的Sheet | 批量文件: 一次上传多个文件合并处理"
    )

# 根据模式渲染不同的 uploader
if processing_mode == "批量文件处理":
    uploaded_files = st.file_uploader(
        "请上传多个Excel文件 (.xlsx)",
        type=['xlsx', 'xls'],
        label_visibility="collapsed",
        accept_multiple_files=True,
        key="batch_uploader",
        help="上传多个包含目标值和计分规则列的Excel文件"
    )
    uploaded_file = None  # 批量模式不使用单文件变量
else:
    uploaded_file = st.file_uploader(
        "请上传Excel文件 (.xlsx)",
        type=['xlsx', 'xls'],
        label_visibility="collapsed",
        key="single_uploader",
        help="上传包含目标值和计分规则列的Excel文件"
    )
    uploaded_files = None  # 非批量模式不使用多文件变量

# 处理按钮 — 批量文件处理模式
if processing_mode == "批量文件处理" and uploaded_files:
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"**已选择 {len(uploaded_files)} 个文件：** " +
                    ", ".join([f"`{f.name}`" for f in uploaded_files]))
    with col2:
        if st.button("🚀 开始批量处理", type="primary", use_container_width=True):
            with st.spinner("正在批量处理数据，请稍候..."):
                try:
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    def progress_callback(progress):
                        progress_bar.progress(progress)
                        status_text.text(f"正在处理... {progress}%")

                    batch_proc = BSCBatchProcessor()
                    files = [(f.name, BytesIO(f.getvalue())) for f in uploaded_files]
                    summary = batch_proc.process(files, progress_callback)

                    # 保存到 session state
                    st.session_state.batch_processor = batch_proc
                    st.session_state.batch_stats = summary
                    st.session_state.is_batch = True
                    st.session_state.is_multi_sheet = False
                    st.session_state.logs = batch_proc.get_logs()

                    # 取第一个成功文件的第一个sheet用于预览
                    if batch_proc.success_files:
                        first_file = batch_proc.success_files[0]
                        first_sheet = list(batch_proc.file_results[first_file].keys())[0]
                        st.session_state.processed_df = batch_proc.file_results[first_file][first_sheet]
                    else:
                        st.session_state.processed_df = None

                    progress_bar.progress(100)
                    status_text.text("✅ 批量处理完成！")

                    if summary['success'] > 0:
                        st.success(f"批量处理完成！成功: {summary['success']}个文件, 失败: {summary['failed']}个文件")
                    else:
                        st.warning(f"所有文件处理失败。失败: {summary['failed']}个文件")

                except Exception as e:
                    st.error(f"批量处理失败：{str(e)}")
                    st.exception(e)

# 处理按钮 — 单文件处理模式（单Sheet / 多Sheet）
elif uploaded_file is not None:
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

                    # 进度条
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    def progress_callback(progress):
                        progress_bar.progress(progress)
                        if progress < 30:
                            status_text.text("正在读取文件...")
                        elif progress < 50:
                            status_text.text("正在识别列...")
                        elif progress < 95:
                            status_text.text("正在处理数据...")
                        else:
                            status_text.text("处理完成！")

                    if processing_mode == "多Sheet处理":
                        # 多Sheet处理模式
                        st.session_state.is_multi_sheet = True
                        st.session_state.is_batch = False
                        multi_processor = BSCMultiSheetProcessor(file_bytes)

                        # 执行处理
                        summary = multi_processor.process(progress_callback)

                        # 保存到session state
                        st.session_state.multi_sheet_processor = multi_processor
                        st.session_state.multi_sheet_stats = summary

                        # 获取第一个成功处理的sheet用于预览
                        if multi_processor.success_sheets:
                            first_sheet = multi_processor.success_sheets[0]
                            st.session_state.processed_df = multi_processor.results[first_sheet]
                        else:
                            st.session_state.processed_df = None

                        st.session_state.logs = multi_processor.get_logs()

                        progress_bar.progress(100)

                        # 显示汇总结果
                        if summary['success'] > 0:
                            st.success(f"多Sheet处理完成！成功: {summary['success']}个, 跳过: {summary['skipped']}个, 失败: {summary['failed']}个")
                        else:
                            st.warning(f"未找到可处理的Sheet。跳过: {summary['skipped']}个, 失败: {summary['failed']}个")

                    else:
                        # 单Sheet处理模式
                        st.session_state.is_multi_sheet = False
                        st.session_state.is_batch = False
                        processor = BSCProcessor(file_bytes)

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

    # 批量文件处理汇总
    if st.session_state.is_batch and st.session_state.batch_stats:
        summary = st.session_state.batch_stats
        batch_proc = st.session_state.batch_processor

        st.markdown("#### 📊 批量文件处理汇总")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("总文件数", summary['total'])
        with col2:
            st.metric("✅ 成功处理", summary['success'], delta_color="normal")
        with col3:
            if summary['failed'] > 0:
                st.metric("❌ 失败", summary['failed'])
            else:
                st.metric("失败", summary['failed'])

        # 显示文件列表
        if summary['success_files']:
            st.markdown("**✅ 成功处理的文件:** " + ", ".join(summary['success_files']))
        if summary['failed_files']:
            st.markdown("**❌ 处理失败的文件:** " + ", ".join(summary['failed_files']))

        st.markdown("---")

        # 两级选择器：先选文件 → 再选Sheet
        if batch_proc and batch_proc.success_files:
            sel_col1, sel_col2 = st.columns(2)
            with sel_col1:
                selected_file = st.selectbox(
                    "选择要预览的文件",
                    batch_proc.success_files,
                    key="batch_file_selector"
                )
            with sel_col2:
                available_sheets = list(batch_proc.file_results[selected_file].keys())
                selected_sheet = st.selectbox(
                    "选择要预览的Sheet",
                    available_sheets,
                    key="batch_sheet_selector"
                )

            st.session_state.processed_df = batch_proc.file_results[selected_file][selected_sheet]

            # 显示该文件的统计信息
            if selected_file in batch_proc.file_stats:
                file_summary = batch_proc.file_stats[selected_file]
                st.markdown(f"**{selected_file}**: 成功 {file_summary.get('success', 0)} 个Sheet, "
                           f"跳过 {file_summary.get('skipped', 0)} 个, "
                           f"失败 {file_summary.get('failed', 0)} 个")

            # 当前选中sheet的半年度统计
            current_df = st.session_state.processed_df
            if current_df is not None and '半年度_解析状态' in current_df.columns:
                semi_counts = current_df['半年度_解析状态'].value_counts()
                st.markdown("#### 半年度处理统计")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("✅ 成功解析", int(semi_counts.get('成功', 0)))
                with col2:
                    st.metric("⚠️ 需人工校验", int(semi_counts.get('人工校验', 0)))
                with col3:
                    st.metric("无半年度数据", int(semi_counts.get('无半年度数据', 0)))
                with col4:
                    st.metric("❌ 错误", int(sum(cnt for status, cnt in semi_counts.items() if 'ERROR' in status)))

    # 多Sheet处理汇总
    elif st.session_state.is_multi_sheet and st.session_state.multi_sheet_stats:
        summary = st.session_state.multi_sheet_stats

        # 显示多Sheet处理汇总
        st.markdown("#### 📊 多Sheet处理汇总")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("总Sheet数", summary['total'])
        with col2:
            st.metric("✅ 成功处理", summary['success'], delta_color="normal")
        with col3:
            st.metric("⚠️ 跳过", summary['skipped'], delta_color="inverse")
        with col4:
            if summary['failed'] > 0:
                st.metric("❌ 失败", summary['failed'])
            else:
                st.metric("失败", summary['failed'])

        # 显示各Sheet列表
        if summary['success_sheets']:
            st.markdown("**✅ 成功处理的Sheet:** " + ", ".join(summary['success_sheets']))
        if summary['skipped_sheets']:
            st.markdown("**⚠️ 跳过的Sheet（无有效列）:** " + ", ".join(summary['skipped_sheets']))
        if summary['failed_sheets']:
            st.markdown("**❌ 处理失败的Sheet:** " + ", ".join(summary['failed_sheets']))

        st.markdown("---")

        # 如果有多个成功处理的sheet，显示sheet选择器
        multi_processor = st.session_state.multi_sheet_processor
        if multi_processor and len(multi_processor.success_sheets) > 1:
            selected_sheet = st.selectbox(
                "选择要预览的Sheet",
                multi_processor.success_sheets,
                key="sheet_selector"
            )
            st.session_state.processed_df = multi_processor.results[selected_sheet]

            # 显示该sheet的统计信息
            if selected_sheet in multi_processor.stats:
                sheet_stats = multi_processor.stats[selected_sheet]
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric(f"{selected_sheet} - 总行数", sheet_stats.get('total', 0))
                with col2:
                    st.metric("成功解析", sheet_stats.get('success', 0))
                with col3:
                    st.metric("人工校验", sheet_stats.get('manual_check', 0))
                with col4:
                    st.metric("错误", sheet_stats.get('error', 0))

                # 半年度统计
                if 'semi_annual' in sheet_stats:
                    semi = sheet_stats['semi_annual']
                    st.markdown("#### 半年度处理统计")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("✅ 成功解析", semi['success'])
                    with col2:
                        st.metric("⚠️ 需人工校验", semi['manual_check'])
                    with col3:
                        st.metric("无半年度数据", semi['no_data'])
                    with col4:
                        st.metric("❌ 错误", semi['error'])
        elif multi_processor and len(multi_processor.success_sheets) == 1:
            # 只有一个成功sheet，直接显示其统计
            only_sheet = multi_processor.success_sheets[0]
            if only_sheet in multi_processor.stats:
                sheet_stats = multi_processor.stats[only_sheet]
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric(f"{only_sheet} - 总行数", sheet_stats.get('total', 0))
                with col2:
                    st.metric("成功解析", sheet_stats.get('success', 0))
                with col3:
                    st.metric("人工校验", sheet_stats.get('manual_check', 0))
                with col4:
                    st.metric("错误", sheet_stats.get('error', 0))

                # 半年度统计
                if 'semi_annual' in sheet_stats:
                    semi = sheet_stats['semi_annual']
                    st.markdown("#### 半年度处理统计")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("✅ 成功解析", semi['success'])
                    with col2:
                        st.metric("⚠️ 需人工校验", semi['manual_check'])
                    with col3:
                        st.metric("无半年度数据", semi['no_data'])
                    with col4:
                        st.metric("❌ 错误", semi['error'])
    else:
        # 单Sheet统计信息
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

            # 半年度统计
            if 'semi_annual' in stats:
                semi = stats['semi_annual']
                st.markdown("#### 半年度处理统计")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("✅ 成功解析", semi['success'])
                with col2:
                    st.metric("⚠️ 需人工校验", semi['manual_check'])
                with col3:
                    st.metric("无半年度数据", semi['no_data'])
                with col4:
                    st.metric("❌ 错误", semi['error'])

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
        annual_status = row.get('解析状态', '')
        semi_status = row.get('半年度_解析状态', '')
        if str(annual_status).startswith('ERROR') or str(semi_status).startswith('ERROR'):
            return ['background-color: #f8d7da'] * len(row)
        elif annual_status == '人工校验' or semi_status == '人工校验':
            return ['background-color: #fff3cd'] * len(row)
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
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"BSC处理结果_{timestamp}.xlsx"

        if st.session_state.is_batch and st.session_state.batch_processor:
            excel_data = st.session_state.batch_processor.save_to_bytesio()
        elif st.session_state.is_multi_sheet and st.session_state.multi_sheet_processor:
            excel_data = st.session_state.multi_sheet_processor.save_to_bytesio()
        else:
            excel_data = st.session_state.processor.save_to_bytesio()

        st.download_button(
            label="📥 下载 Excel 文件",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        # 同时提供CSV下载选项（仅当前预览的sheet）
        csv_data = df[display_columns].to_csv(index=False, encoding='utf-8-sig')
        csv_filename = f"BSC处理结果_{timestamp}.csv"

        st.download_button(
            label="📄 下载 CSV 文件",
            data=csv_data,
            file_name=csv_filename,
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
if uploaded_file is None and not uploaded_files and st.session_state.processed_df is None:
    st.markdown("---")
    st.markdown("""
    ### 👋 欢迎使用平衡计分卡KPI数据处理器

    请在上方上传您的Excel文件开始处理。

    **文件要求：**
    - Excel文件需包含 **目标值列**（列名包含"目标值"关键字）
    - Excel文件需包含 **计分规则列**（列名包含"计分规则"关键字）

    如有问题，请查看左侧功能说明。
    """)
