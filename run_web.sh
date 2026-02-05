#!/bin/bash
# Streamlit Web 应用启动脚本

echo "正在启动平衡计分卡 KPI 数据处理器..."
echo ""

# 检查是否安装了streamlit
if ! python3 -c "import streamlit" 2>/dev/null; then
    echo "错误: 未安装 streamlit"
    echo "请运行: pip install streamlit"
    exit 1
fi

# 启动应用
streamlit run bsc_web.py --server.port 8501 --server.address localhost
