"""
Excel解析模块
用于读取和解析Excel中的信访登记数据
"""
import pandas as pd
import io
import msoffcrypto
from date_calculator import (
    parse_excel_date, 
    get_current_week_range, 
    get_last_week_range,
    is_date_in_range
)


class ExcelParser:
    """Excel数据解析器"""
    
    def __init__(self, excel_path, password=None):
        """
        初始化Excel解析器
        
        Args:
            excel_path: Excel文件路径
            password: Excel文件密码（如果文件有密码保护）
        """
        self.excel_path = excel_path
        self.password = password
        self.sunshine_sheet_name = "阳光xf登记"
        self.gab_sheet_name = "gab上访"
        self.decrypted_file = None
        
        # 获取本周和上周的日期范围
        self.current_week_start, self.current_week_end = get_current_week_range()
        self.last_week_start, self.last_week_end = get_last_week_range()
        
        # 如果提供了密码，先解密文件
        if self.password:
            self._decrypt_file()
    
    def _decrypt_file(self):
        """解密Excel文件"""
        try:
            with open(self.excel_path, 'rb') as f:
                file = msoffcrypto.OfficeFile(f)
                file.load_key(password=self.password)
                
                # 将解密后的内容存储在内存中
                self.decrypted_file = io.BytesIO()
                file.decrypt(self.decrypted_file)
                self.decrypted_file.seek(0)
        except Exception as e:
            print(f"解密文件失败: {e}")
            self.decrypted_file = None
    
    def parse_sheet(self, sheet_name):
        """
        解析指定sheet的数据
        
        Args:
            sheet_name: sheet名称
        
        Returns:
            dict: 包含本周和上周人数的字典
        """
        try:
            # 读取Excel的指定sheet
            # 如果文件已解密，使用解密后的内容；否则直接读取文件
            excel_source = self.decrypted_file if self.decrypted_file else self.excel_path
            
            # 自动检测文件格式并使用相应引擎
            # .xls使用xlrd引擎，.xlsx使用openpyxl引擎
            if self.excel_path.endswith('.xls'):
                df = pd.read_excel(excel_source, sheet_name=sheet_name, engine='xlrd')
            else:
                df = pd.read_excel(excel_source, sheet_name=sheet_name, engine='openpyxl')
            
            # 检查是否存在"登记时间"列（B列，索引1）
            if len(df.columns) < 2:
                return {'current_week': 0, 'last_week': 0, 'error': '表格格式不正确'}
            
            # 获取B列（登记时间列）
            # 列名可能是"登记时间"或者是列索引
            date_column = df.iloc[:, 1]  # 第2列（B列）
            
            current_week_count = 0
            last_week_count = 0
            
            # 遍历每一行
            for idx, date_value in enumerate(date_column):
                # 跳过表头行
                if idx == 0:
                    continue
                
                # 解析日期
                parsed_date = parse_excel_date(date_value)
                
                if parsed_date:
                    # 判断是否在本周
                    if is_date_in_range(parsed_date, self.current_week_start, self.current_week_end):
                        current_week_count += 1
                    # 判断是否在上周
                    elif is_date_in_range(parsed_date, self.last_week_start, self.last_week_end):
                        last_week_count += 1
            
            return {
                'current_week': current_week_count,
                'last_week': last_week_count
            }
            
        except Exception as e:
            return {
                'current_week': 0,
                'last_week': 0,
                'error': str(e)
            }
    
    def parse_all(self):
        """
        解析所有sheet的数据
        
        Returns:
            dict: 包含所有统计数据的字典
        """
        # 解析"阳光xf登记"sheet
        sunshine_data = self.parse_sheet(self.sunshine_sheet_name)
        
        # 解析"gab上访"sheet
        gab_data = self.parse_sheet(self.gab_sheet_name)
        
        # 汇总数据
        result = {
            # 阳光xf登记数据
            'sunshine_current': sunshine_data['current_week'],
            'sunshine_last': sunshine_data['last_week'],
            
            # gab上访数据
            'gab_current': gab_data['current_week'],
            'gab_last': gab_data['last_week'],
            
            # 总计
            'total_current': sunshine_data['current_week'] + gab_data['current_week'],
            
            # 环比趋势
            'sunshine_trend': self._calculate_trend(
                sunshine_data['current_week'], 
                sunshine_data['last_week']
            ),
            'gab_trend': self._calculate_trend(
                gab_data['current_week'], 
                gab_data['last_week']
            ),
            
            # 错误信息
            'errors': []
        }
        
        # 收集错误信息
        if 'error' in sunshine_data:
            result['errors'].append(f"阳光xf登记: {sunshine_data['error']}")
        if 'error' in gab_data:
            result['errors'].append(f"gab上访: {gab_data['error']}")
        
        return result
    
    def _calculate_trend(self, current, last):
        """
        计算环比趋势
        
        Args:
            current: 本周数量
            last: 上周数量
        
        Returns:
            str: 趋势描述文本
        """
        diff = current - last
        
        if diff > 0:
            return f"上升{diff}人"
        elif diff < 0:
            return f"下降{abs(diff)}人"
        else:
            return "持平"


def test_parser():
    """测试Excel解析器"""
    print("=== Excel解析模块测试 ===\n")
    
    # 测试文件路径
    test_file = "2025年复盘人员明细9.22.xls"
    password = "110110"  # 文件密码
    
    try:
        parser = ExcelParser(test_file, password=password)
        result = parser.parse_all()
        
        print(f"本周日期范围: {parser.current_week_start.strftime('%Y-%m-%d')} 到 {parser.current_week_end.strftime('%Y-%m-%d')}")
        print(f"上周日期范围: {parser.last_week_start.strftime('%Y-%m-%d')} 到 {parser.last_week_end.strftime('%Y-%m-%d')}\n")
        
        print("统计结果:")
        print(f"  阳光xf登记 - 本周: {result['sunshine_current']}人, 上周: {result['sunshine_last']}人, 环比: {result['sunshine_trend']}")
        print(f"  gab上访 - 本周: {result['gab_current']}人, 上周: {result['gab_last']}人, 环比: {result['gab_trend']}")
        print(f"  总计 - 本周: {result['total_current']}人")
        
        if result['errors']:
            print("\n错误信息:")
            for error in result['errors']:
                print(f"  - {error}")
    
    except FileNotFoundError:
        print(f"❌ 测试文件 {test_file} 不存在")
    except Exception as e:
        print(f"❌ 解析失败: {e}")


if __name__ == '__main__':
    test_parser()

