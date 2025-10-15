"""
日期计算模块
用于计算本周和上周的日期范围（周一到周日）
"""
from datetime import datetime, timedelta


def get_week_range(date=None):
    """
    获取指定日期所在周的周一和周日
    
    Args:
        date: datetime对象，默认为当前日期
    
    Returns:
        tuple: (周一, 周日) 的datetime对象
    """
    if date is None:
        date = datetime.now()
    
    # 获取当前是星期几（0=Monday, 6=Sunday）
    weekday = date.weekday()
    
    # 计算本周一
    monday = date - timedelta(days=weekday)
    # 计算本周日
    sunday = monday + timedelta(days=6)
    
    # 只保留日期部分，去除时间
    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
    sunday = sunday.replace(hour=23, minute=59, second=59, microsecond=999999)
    
    return monday, sunday


def get_current_week_range():
    """
    获取本周的日期范围（周一到周日）
    
    Returns:
        tuple: (本周一, 本周日) 的datetime对象
    """
    return get_week_range()


def get_last_week_range():
    """
    获取上周的日期范围（上周一到上周日）
    
    Returns:
        tuple: (上周一, 上周日) 的datetime对象
    """
    # 获取上周的任意一天（本周一减7天）
    this_monday, _ = get_current_week_range()
    last_week_date = this_monday - timedelta(days=7)
    
    return get_week_range(last_week_date)


def is_date_in_range(date, start_date, end_date):
    """
    判断日期是否在指定范围内
    
    Args:
        date: 要判断的日期
        start_date: 开始日期
        end_date: 结束日期
    
    Returns:
        bool: 是否在范围内
    """
    return start_date <= date <= end_date


def parse_excel_date(date_str, year=2025):
    """
    解析Excel中的日期格式（月.日）为完整日期
    
    Args:
        date_str: 日期字符串，格式为 "月.日"，例如 "1.2", "2.5"
        year: 年份，默认为2025
    
    Returns:
        datetime: 解析后的日期对象，如果解析失败返回None
    """
    try:
        # 处理可能的空值或NaN
        if not date_str or str(date_str).strip() == '' or str(date_str).lower() == 'nan':
            return None
            
        date_str = str(date_str).strip()
        
        # 解析 "月.日" 格式
        parts = date_str.split('.')
        if len(parts) != 2:
            return None
        
        month = int(parts[0])
        day = int(parts[1])
        
        # 创建日期对象
        return datetime(year, month, day)
    except (ValueError, AttributeError):
        return None


if __name__ == '__main__':
    # 测试代码
    print("=== 日期计算模块测试 ===\n")
    
    # 测试本周范围
    current_monday, current_sunday = get_current_week_range()
    print(f"本周范围: {current_monday.strftime('%Y-%m-%d')} 到 {current_sunday.strftime('%Y-%m-%d')}")
    
    # 测试上周范围
    last_monday, last_sunday = get_last_week_range()
    print(f"上周范围: {last_monday.strftime('%Y-%m-%d')} 到 {last_sunday.strftime('%Y-%m-%d')}")
    
    # 测试日期解析
    print("\n=== 日期解析测试 ===")
    test_dates = ["1.2", "1.6", "2.5", "1.17", "2.26"]
    for date_str in test_dates:
        parsed = parse_excel_date(date_str)
        if parsed:
            print(f"{date_str} -> {parsed.strftime('%Y-%m-%d')}")
    
    # 测试日期范围判断
    print("\n=== 日期范围判断测试 ===")
    test_date = parse_excel_date("1.17")
    if test_date:
        in_current = is_date_in_range(test_date, current_monday, current_sunday)
        in_last = is_date_in_range(test_date, last_monday, last_sunday)
        print(f"1.17 在本周: {in_current}, 在上周: {in_last}")


