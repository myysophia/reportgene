#!/usr/bin/env python3
"""
系统验证脚本
检查所有组件是否正常工作
"""
import sys
import os

def check_imports():
    """检查所有必要的模块导入"""
    print("🔍 检查模块导入...")
    
    try:
        import gradio
        print(f"  ✓ Gradio {gradio.__version__}")
    except ImportError as e:
        print(f"  ❌ Gradio导入失败: {e}")
        return False
    
    try:
        import pandas
        print(f"  ✓ Pandas {pandas.__version__}")
    except ImportError as e:
        print(f"  ❌ Pandas导入失败: {e}")
        return False
    
    try:
        import openpyxl
        print(f"  ✓ OpenPyXL {openpyxl.__version__}")
    except ImportError as e:
        print(f"  ❌ OpenPyXL导入失败: {e}")
        return False
    
    try:
        import docx
        print(f"  ✓ python-docx")
    except ImportError as e:
        print(f"  ❌ python-docx导入失败: {e}")
        return False
    
    try:
        import msoffcrypto
        print(f"  ✓ msoffcrypto-tool")
    except ImportError as e:
        print(f"  ❌ msoffcrypto-tool导入失败: {e}")
        return False
    
    try:
        import audioop
        print(f"  ✓ audioop (通过audioop-lts)")
    except ImportError as e:
        print(f"  ❌ audioop导入失败: {e}")
        return False
    
    return True

def check_files():
    """检查必要文件是否存在"""
    print("\n📁 检查文件...")
    
    required_files = [
        "app.py",
        "excel_parser.py", 
        "date_calculator.py",
        "word_generator.py",
        "template.docx",
        "requirements.txt",
        "start.sh"
    ]
    
    all_exist = True
    for file in required_files:
        if os.path.exists(file):
            print(f"  ✓ {file}")
        else:
            print(f"  ❌ {file} 不存在")
            all_exist = False
    
    return all_exist

def check_modules():
    """检查自定义模块"""
    print("\n🔧 检查自定义模块...")
    
    try:
        from date_calculator import get_current_week_range, parse_excel_date
        print("  ✓ date_calculator 模块")
    except ImportError as e:
        print(f"  ❌ date_calculator 模块: {e}")
        return False
    
    try:
        from excel_parser import ExcelParser
        print("  ✓ excel_parser 模块")
    except ImportError as e:
        print(f"  ❌ excel_parser 模块: {e}")
        return False
    
    try:
        from word_generator import WordGenerator
        print("  ✓ word_generator 模块")
    except ImportError as e:
        print(f"  ❌ word_generator 模块: {e}")
        return False
    
    try:
        import app
        print("  ✓ app 模块")
    except ImportError as e:
        print(f"  ❌ app 模块: {e}")
        return False
    
    return True

def check_directories():
    """检查必要目录"""
    print("\n📂 检查目录...")
    
    directories = ["output", "upload"]
    all_exist = True
    
    for dir_name in directories:
        if not os.path.exists(dir_name):
            print(f"  ⚠️  {dir_name}目录不存在，正在创建...")
            os.makedirs(dir_name, exist_ok=True)
        
        if os.path.exists(dir_name):
            print(f"  ✓ {dir_name}目录存在")
        else:
            print(f"  ❌ 无法创建{dir_name}目录")
            all_exist = False
    
    return all_exist

def main():
    """主验证函数"""
    print("=" * 60)
    print("🧪 汇享易报告生成系统 - 系统验证")
    print("=" * 60)
    
    checks = [
        ("模块导入", check_imports),
        ("文件检查", check_files),
        ("自定义模块", check_modules),
        ("目录检查", check_directories)
    ]
    
    all_passed = True
    
    for name, check_func in checks:
        if not check_func():
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("✅ 系统验证通过！所有组件正常")
        print("\n🚀 可以启动系统：")
        print("   ./start.sh")
        print("\n🌐 访问地址：")
        print("   http://localhost:7861")
    else:
        print("❌ 系统验证失败！请检查上述错误")
        print("\n🔧 建议操作：")
        print("   1. 重新安装依赖：pip install -r requirements.txt")
        print("   2. 检查Python版本：python --version")
        print("   3. 查看错误信息并修复")
    
    print("=" * 60)
    
    return 0 if all_passed else 1

if __name__ == "__main__":
    sys.exit(main())
