#!/bin/bash
# Docker快速启动脚本

set -e

echo "🐳 汇享易报告自助生成智能体 - Docker部署"
echo "=========================================="
echo ""

# 检查Docker是否安装
if ! command -v docker &> /dev/null; then
    echo "❌ Docker未安装，请先安装Docker"
    echo "   访问：https://docs.docker.com/get-docker/"
    exit 1
fi

# 检查Docker Compose是否安装
if ! command -v docker-compose &> /dev/null && ! docker compose version &> /dev/null; then
    echo "❌ Docker Compose未安装，请先安装Docker Compose"
    exit 1
fi

# 检查必要文件
if [ ! -f "template.docx" ]; then
    echo "❌ 模板文件 template.docx 不存在"
    exit 1
fi

echo "✓ Docker环境检查通过"
echo ""

# 创建必要目录
mkdir -p upload output
echo "✓ 创建目录：upload/ output/"
echo ""

# 构建并启动
echo "🔨 构建Docker镜像..."
docker-compose build

echo ""
echo "🚀 启动容器..."
docker-compose up -d

echo ""
echo "⏳ 等待服务启动..."
sleep 5

# 检查容器状态
if docker-compose ps | grep -q "Up"; then
    echo ""
    echo "=========================================="
    echo "✅ 服务启动成功！"
    echo "=========================================="
    echo ""
    echo "📱 访问地址："
    echo "   http://localhost:7861"
    echo ""
    echo "📋 常用命令："
    echo "   查看日志：docker-compose logs -f"
    echo "   停止服务：docker-compose down"
    echo "   重启服务：docker-compose restart"
    echo ""
    echo "📚 详细文档："
    echo "   Docker使用说明.md"
    echo ""
    echo "=========================================="
else
    echo ""
    echo "❌ 服务启动失败，请查看日志："
    echo "   docker-compose logs"
    exit 1
fi

