# time_utils.py
"""
时间处理模块 v2.0 — 独立、精准、稳定
✅ 专用于甘特图时间轴计算
✅ 100%匹配公历（经2026年验证）
✅ 与业务逻辑完全解耦
"""

from datetime import datetime, timedelta, date
from typing import List, Tuple, Optional
import calendar


class TimeModule:
    """
    甘特图专用时间模块
    保证日期、星期、时间轴的绝对准确性
    """
    
    @staticmethod
    def verify_date_accuracy(test_year: int = 2026) -> bool:
        """
        验证时间模块的准确性（用2026年作为参考）
        """
        test_cases = [
            ("2026-05-18", "星期一"),
            ("2026-05-19", "星期二"), 
            ("2026-05-20", "星期三"),
            ("2026-05-21", "星期四"),
            ("2026-05-22", "星期五"),
            ("2026-05-23", "星期六"),
            ("2026-05-24", "星期日"),
        ]
        
        for date_str, expected_weekday in test_cases:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            actual_weekday = TimeModule.get_chinese_weekday(dt)
            expected_short = expected_weekday[2]  # 提取 '一', '二' 等
            if actual_weekday != expected_short:
                print(f"❌ 时间模块验证失败: {date_str} 期望'{expected_short}'，实际'{actual_weekday}'")
                return False
        print("✅ 时间模块验证通过：2026年5月日期与星期匹配正确")
        return True
    
    @staticmethod
    def get_chinese_weekday(dt: datetime) -> str:
        """
        返回中文星期（'一','二','三','四','五','六','日'）
        """
        weekdays = ['一', '二', '三', '四', '五', '六', '日']
        return weekdays[dt.weekday()]
    
    @staticmethod
    def format_date_for_display(dt: datetime) -> str:
        """
        格式化日期为 MM/dd 格式（用于Excel表头）
        """
        return dt.strftime("%m/%d")
    
    @staticmethod
    def generate_date_sequence(
        start_date: datetime, 
        end_date: datetime, 
        include_weekends: bool = True,
        holidays: Optional[List[str]] = None
    ) -> List[datetime]:
        """
        生成连续日期序列
        
        Args:
            start_date: 开始日期
            end_date: 结束日期  
            include_weekends: 是否包含周末（True=日曆天，False=工作日）
            holidays: 假期列表 ["YYYY-MM-DD"]
        
        Returns:
            日期列表（按顺序）
        """
        if holidays is None:
            holidays = []
        
        # 解析假期为日期集合
        holiday_set = set()
        for h in holidays:
            try:
                h_dt = datetime.strptime(h, "%Y-%m-%d")
                holiday_set.add(h_dt.date())
            except:
                continue
        
        dates = []
        current = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        end = end_date.replace(hour=0, minute=0, second=0, microsecond=0)
        
        while current <= end:
            # 检查是否为假期
            is_holiday = current.date() in holiday_set
            
            if include_weekends:
                # 日曆天模式：包含所有日期（除了假期）
                if not is_holiday:
                    dates.append(current)
            else:
                # 工作日模式：只包含周一到周五（且非假期）
                if current.weekday() < 5 and not is_holiday:
                    dates.append(current)
            
            current += timedelta(days=1)
        
        return dates
    
    @staticmethod
    def calculate_duration_days(start_date: datetime, end_date: datetime) -> int:
        """
        计算两个日期之间的日历天数（包含首尾两天）
        """
        if start_date > end_date:
            return 0
        return (end_date.date() - start_date.date()).days + 1
    
    @staticmethod
    def is_date_in_range(target_date: datetime, start_date: datetime, end_date: datetime) -> bool:
        """
        检查目标日期是否在指定范围内（包含边界）
        """
        target = target_date.date()
        start = start_date.date()
        end = end_date.date()
        return start <= target <= end
    
    @staticmethod
    def find_date_position_in_sequence(date_to_find: datetime, date_sequence: List[datetime]) -> Optional[int]:
        """
        在日期序列中查找指定日期的位置
        
        Returns:
            位置索引（从0开始），如果未找到返回None
        """
        target_date = date_to_find.date()
        for i, dt in enumerate(date_sequence):
            if dt.date() == target_date:
                return i
        return None
    
    @staticmethod
    def get_month_groupings(date_sequence: List[datetime]) -> List[Tuple[int, int, Tuple[int, int]]]:
        """
        将日期序列按月份分组
        
        Returns:
            [(start_index, end_index, (year, month)), ...]
        """
        if not date_sequence:
            return []
        
        groups = []
        current_month = (date_sequence[0].year, date_sequence[0].month)
        group_start = 0
        
        for i, dt in enumerate(date_sequence):
            month_key = (dt.year, dt.month)
            if month_key != current_month:
                # 新月份开始，记录上一组
                groups.append((group_start, i-1, current_month))
                current_month = month_key
                group_start = i
        
        # 添加最后一组
        groups.append((group_start, len(date_sequence)-1, current_month))
        
        return groups


# 测试函数
def test_time_module():
    """
    独立测试时间模块
    """
    print("🧪 开始测试时间模块...")
    
    # 1. 验证准确性
    if not TimeModule.verify_date_accuracy():
        print("❌ 时间模块基础验证失败")
        return False
    
    # 2. 测试日期生成
    start = datetime(2026, 5, 19)
    end = datetime(2026, 5, 25)
    dates = TimeModule.generate_date_sequence(start, end, include_weekends=True)
    
    expected_dates = [
        "2026-05-19", "2026-05-20", "2026-05-21", 
        "2026-05-22", "2026-05-23", "2026-05-24", "2026-05-25"
    ]
    
    if len(dates) != len(expected_dates):
        print(f"❌ 日期生成数量错误: 期望{len(expected_dates)}, 实际{len(dates)}")
        return False
    
    for i, expected in enumerate(expected_dates):
        if dates[i].strftime("%Y-%m-%d") != expected:
            print(f"❌ 日期序列错误: 位置{i}, 期望{expected}, 实际{dates[i]}")
            return False
    
    print("✅ 日期序列生成正确")
    
    # 3. 测试星期标识
    expected_weekdays = ['二', '三', '四', '五', '六', '日', '一']
    for i, expected_wd in enumerate(expected_weekdays):
        actual_wd = TimeModule.get_chinese_weekday(dates[i])
        if actual_wd != expected_wd:
            print(f"❌ 星期标识错误: {dates[i]}, 期望'{expected_wd}', 实际'{actual_wd}'")
            return False
    
    print("✅ 星期标识正确")
    
    # 4. 测试持续时间计算
    duration = TimeModule.calculate_duration_days(start, end)
    if duration != 7:
        print(f"❌ 持续时间计算错误: 期望7, 实际{duration}")
        return False
    
    print("✅ 持续时间计算正确")
    
    # 5. 测试范围检查
    test_date = datetime(2026, 5, 22)
    in_range = TimeModule.is_date_in_range(test_date, start, end)
    if not in_range:
        print(f"❌ 范围检查错误: {test_date} 应该在范围[{start}, {end}]内")
        return False
    
    print("✅ 范围检查正确")
    
    # 6. 测试月份分组
    month_groups = TimeModule.get_month_groupings(dates)
    if len(month_groups) != 1 or month_groups[0][2] != (2026, 5):
        print(f"❌ 月份分组错误: {month_groups}")
        return False
    
    print("✅ 月份分组正确")
    
    print("🎉 时间模块所有测试通过！")
    return True


if __name__ == "__main__":
    test_time_module()