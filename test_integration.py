"""
集成测试脚本
测试完整的报告生成流程
"""
import os
from excel_parser import ExcelParser
from word_generator import WordGenerator
from date_calculator import get_current_week_range, get_last_week_range


def test_full_workflow():
    """测试完整工作流程"""
    print("=" * 60)
    print("🧪 汇享易报告生成系统 - 集成测试")
    print("=" * 60)
    print()
    
    # 配置
    excel_file = "2025年复盘人员明细9.22.xls"
    excel_password = "110110"
    template_file = "template.docx"
    output_file = "output/集成测试报告.docx"
    
    # 步骤1: 检查文件
    print("📋 步骤1: 检查文件...")
    print(f"  - Excel文件: {excel_file}")
    print(f"  - 模板文件: {template_file}")
    
    if not os.path.exists(excel_file):
        print(f"  ❌ Excel文件不存在: {excel_file}")
        return False
    
    if not os.path.exists(template_file):
        print(f"  ❌ 模板文件不存在: {template_file}")
        return False
    
    print("  ✅ 文件检查通过\n")
    
    # 步骤2: 解析Excel
    print("📊 步骤2: 解析Excel数据...")
    try:
        parser = ExcelParser(excel_file, password=excel_password)
        
        # 显示日期范围
        current_start, current_end = get_current_week_range()
        last_start, last_end = get_last_week_range()
        
        print(f"  - 本周范围: {current_start.strftime('%Y-%m-%d')} 到 {current_end.strftime('%Y-%m-%d')}")
        print(f"  - 上周范围: {last_start.strftime('%Y-%m-%d')} 到 {last_end.strftime('%Y-%m-%d')}")
        
        # 解析数据
        data = parser.parse_all()
        
        if data.get('errors'):
            print("  ⚠️  解析过程中遇到问题:")
            for error in data['errors']:
                print(f"    - {error}")
        
        print(f"\n  📈 统计结果:")
        print(f"    - 阳光xf登记: 本周 {data['sunshine_current']} 人, 上周 {data['sunshine_last']} 人, {data['sunshine_trend']}")
        print(f"    - gab上访: 本周 {data['gab_current']} 人, 上周 {data['gab_last']} 人, {data['gab_trend']}")
        print(f"    - 本周总计: {data['total_current']} 人")
        print("  ✅ Excel解析成功\n")
        
    except Exception as e:
        print(f"  ❌ Excel解析失败: {e}\n")
        return False
    
    # 步骤3: 生成Word
    print("📝 步骤3: 生成Word文档...")
    try:
        generator = WordGenerator(template_file)
        success = generator.generate(data, output_file)
        
        if success:
            print(f"  ✅ Word文档生成成功: {output_file}")
            print(f"  📄 文件大小: {os.path.getsize(output_file)} 字节\n")
        else:
            print(f"  ❌ Word文档生成失败\n")
            return False
            
    except Exception as e:
        print(f"  ❌ Word生成失败: {e}\n")
        return False
    
    # 步骤4: 验证输出
    print("✅ 步骤4: 验证输出文件...")
    if os.path.exists(output_file):
        print(f"  ✅ 输出文件存在: {output_file}")
        print(f"  ✅ 文件可访问\n")
    else:
        print(f"  ❌ 输出文件不存在\n")
        return False
    
    # 总结
    print("=" * 60)
    print("✅ 集成测试完成！所有步骤通过")
    print("=" * 60)
    print()
    print("📊 最终数据汇总:")
    print(f"  • 本周信访登记总人数: {data['total_current']} 人")
    print(f"  • 阳光xf登记: {data['sunshine_current']} 人 ({data['sunshine_trend']})")
    print(f"  • gab上访: {data['gab_current']} 人 ({data['gab_trend']})")
    print()
    print(f"📄 报告已生成: {output_file}")
    print()
    
    return True


if __name__ == '__main__':
    success = test_full_workflow()
    exit(0 if success else 1)


