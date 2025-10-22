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
            dict: 包含本周和上周人数及详细信息的字典
        """
        try:
            # 读取Excel的指定sheet（不使用header，原始读取）
            excel_source = self.decrypted_file if self.decrypted_file else self.excel_path
            
            if self.excel_path.endswith('.xls'):
                df = pd.read_excel(excel_source, sheet_name=sheet_name, engine='xlrd', header=None)
            else:
                df = pd.read_excel(excel_source, sheet_name=sheet_name, engine='openpyxl', header=None)
            
            # 检查表格格式
            if len(df.columns) < 2:
                return {'current_week': 0, 'last_week': 0, 'error': '表格格式不正确'}
            
            current_week_count = 0
            last_week_count = 0
            current_week_persons = []  # 本周人员详细信息
            
            # 从第3行开始遍历数据（跳过标题和空行，索引从0开始，实际第3行是索引2）
            for idx in range(2, len(df)):
                date_value = df.iloc[idx, 1]  # B列（登记时间）
                
                # 解析日期
                parsed_date = parse_excel_date(date_value)
                
                if parsed_date:
                    # 提取行数据
                    unit = df.iloc[idx, 9] if pd.notna(df.iloc[idx, 9]) else ""  # 责任单位（J列，索引9）
                    name = df.iloc[idx, 2] if pd.notna(df.iloc[idx, 2]) else "XX"  # 姓名（C列，索引2）
                    travel_method = df.iloc[idx, 16] if pd.notna(df.iloc[idx, 16]) else ""  # 进京方式（Q列，索引16）
                    group_appeal = df.iloc[idx, 18] if pd.notna(df.iloc[idx, 18]) else ""  # 群体诉求（S列，索引18）
                    
                    # 判断是否在本周
                    if is_date_in_range(parsed_date, self.current_week_start, self.current_week_end):
                        current_week_count += 1
                        current_week_persons.append({
                            'unit': str(unit),
                            'name': str(name),
                            'travel_method': str(travel_method),
                            'group_appeal': str(group_appeal)
                        })
                    # 判断是否在上周
                    elif is_date_in_range(parsed_date, self.last_week_start, self.last_week_end):
                        last_week_count += 1
            
            return {
                'current_week': current_week_count,
                'last_week': last_week_count,
                'persons': current_week_persons
            }
            
        except Exception as e:
            return {
                'current_week': 0,
                'last_week': 0,
                'persons': [],
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
        
        # 合并本周所有人员
        all_persons = sunshine_data.get('persons', []) + gab_data.get('persons', [])
        
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
            
            # 格式化的人员信息
            'sunshine_persons_text': self._format_persons_list(sunshine_data.get('persons', [])),
            'gab_persons_text': self._format_persons_list(gab_data.get('persons', [])),
            
            # 地区统计
            'area_stats_text': self._format_area_stats(all_persons),
            
            # 群体诉求统计
            'group_appeal_text': self._format_group_appeal_stats(all_persons),
            
            # 进京方式统计
            'travel_road_count': self._count_travel_method(all_persons, '公路'),
            'travel_stats_text': self._format_travel_stats(all_persons),
            
            # 错误信息
            'errors': []
        }
        
        # 收集错误信息
        if 'error' in sunshine_data:
            result['errors'].append(f"阳光xf登记: {sunshine_data['error']}")
        if 'error' in gab_data:
            result['errors'].append(f"gab上访: {gab_data['error']}")
        
        return result
    
    def _format_persons_list(self, persons):
        """格式化人员列表为：单位+姓名 格式"""
        if not persons:
            return ""
        
        formatted = []
        for p in persons:
            unit = p.get('unit', '')
            name = p.get('name', 'XX')
            # 隐藏姓名的最后一个字
            if len(name) > 1:
                name = name[:-1] + 'X'
            formatted.append(f"{unit}{name}")
        
        return "、".join(formatted)
    
    def _format_area_stats(self, persons):
        """统计各地区人数"""
        from collections import Counter
        
        area_counter = Counter()
        for p in persons:
            unit = p.get('unit', '')
            if unit:
                area_counter[unit] += 1
        
        # 格式化输出
        stats_parts = []
        for area, count in sorted(area_counter.items(), key=lambda x: -x[1]):
            if count > 1:
                stats_parts.append(f"{area}{count}人")
            else:
                stats_parts.append(f"{area}1人")
        
        return "，".join(stats_parts) if stats_parts else ""
    
    def _format_group_appeal_stats(self, persons):
        """统计群体诉求类型及人数"""
        from collections import Counter
        
        appeal_counter = Counter()
        for p in persons:
            appeal = p.get('group_appeal', '').strip()
            if appeal and appeal not in ['nan', '']:
                appeal_counter[appeal] += 1
        
        # 格式化输出
        stats_parts = []
        for appeal, count in sorted(appeal_counter.items(), key=lambda x: -x[1]):
            stats_parts.append(f"{appeal}{count}人")
        
        return "、".join(stats_parts) if stats_parts else "无"
    
    def _count_travel_method(self, persons, method):
        """统计特定进京方式的人数"""
        count = 0
        for p in persons:
            travel = p.get('travel_method', '').strip()
            if method in travel:
                count += 1
        return count
    
    def _format_travel_stats(self, persons):
        """统计所有进京方式及人数"""
        from collections import Counter
        
        travel_counter = Counter()
        for p in persons:
            travel = p.get('travel_method', '').strip()
            if travel and travel not in ['nan', '']:
                travel_counter[travel] += 1
        
        # 格式化输出
        stats_parts = []
        for travel, count in sorted(travel_counter.items(), key=lambda x: -x[1]):
            stats_parts.append(f"{travel}{count}人")
        
        return "、".join(stats_parts) if stats_parts else "无"
    
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

