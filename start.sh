#!/bin/bash
# 汇享易报告自助生成智能体启动脚本

echo "🚀 启动汇享易报告自助生成智能体..."
echo ""

# 激活虚拟环境
if [ -d "venv" ]; then
    echo "✓ 激活虚拟环境..."
    source venv/bin/activate
else
    echo "❌ 虚拟环境不存在，请先运行: python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt"
    exit 1
fi

# 检查依赖
echo "✓ 检查依赖..."
python -c "import gradio, pandas, docx" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  依赖缺失，正在安装..."
    pip install -r requirements.txt -q
fi

# 启动应用
echo "✓ 启动应用..."
echo ""
echo "================================================"
echo "  汇享易报告自助生成智能体"
echo "  访问地址: http://localhost:7861"
echo "================================================"
echo ""

python app.py
