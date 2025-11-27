import unittest
from datetime import datetime, date
from input_handler import 我的输入器  # 替换为实际的模块和类名


class TestDateParser(unittest.TestCase):

    def setUp(self):
        """测试前置设置"""
        self.parser = 我的输入器()
        self.today = datetime.now().date()

    def test_standard_date_formats(self):
        """测试标准日期格式"""
        test_cases = [
            # (输入字符串, 期望日期)
            ("2020-7-1", date(2020, 7, 1)),
            ("2020/7/1", date(2020, 7, 1)),
            ("2020 7 1", date(2020, 7, 1)),
            ("2020-07-01", date(2020, 7, 1)),
            ("2020/07/01", date(2020, 7, 1)),
            ("2020 07 01", date(2020, 7, 1)),
        ]

        for input_str, expected_date in test_cases:
            with self.subTest(input=input_str):
                year, month, day = self.parser._get_ymd_time_by_str_save(input_str)
                result_date = date(year, month, day)
                self.assertEqual(result_date, expected_date)

    def test_today_relative_dates(self):
        """测试相对今天的日期"""
        base_date = datetime(2023, 10, 15)  # 固定基准日期

        test_cases = [
            # (输入字符串, 期望日期)
            ("today", date(2023, 10, 15)),
            ("td", date(2023, 10, 15)),
            ("today + 1", date(2023, 10, 16)),
            ("td + 1", date(2023, 10, 16)),
            ("today + 0", date(2023, 10, 15)),
            ("td + 0", date(2023, 10, 15)),
            ("today + 10", date(2023, 10, 25)),
            ("today - 1", date(2023, 10, 14)),
            ("td - 1", date(2023, 10, 14)),
            ("today - 10", date(2023, 10, 5)),
            ("today+1", date(2023, 10, 16)),  # 无空格
            ("td-1", date(2023, 10, 14)),  # 无空格
        ]

        for input_str, expected_date in test_cases:
            with self.subTest(input=input_str):
                year, month, day = self.parser._get_ymd_time_by_str_save(input_str, base_date)
                result_date = date(year, month, day)
                self.assertEqual(result_date, expected_date)

    def test_week_relative_dates(self):
        """测试相对周的日期"""
        # 以2023-10-15（周日）为基准
        base_date = datetime(2023, 10, 15)

        test_cases = [
            # (输入字符串, 期望日期)
            ("week 0", date(2023, 10, 15)),  # 周日
            ("w 0", date(2023, 10, 15)),  # 周日
            ("week 1", date(2023, 10, 9)),  # 周一
            ("w 1", date(2023, 10, 9)),  # 周一
            ("week 7", date(2023, 10, 15)),  # 周日
            ("w 7", date(2023, 10, 15)),  # 周日
            ("week +1 0", date(2023, 10, 22)),  # 下一周周日
            ("w +1 0", date(2023, 10, 22)),  # 下一周周日
            ("week -1 0", date(2023, 10, 8)),  # 上一周周日
            ("w -1 0", date(2023, 10, 8)),  # 上一周周日
            ("week+1 1", date(2023, 10, 16)),  # 下一周周一
            ("w-1 1", date(2023, 10, 2)),  # 上一周周一
        ]

        for input_str, expected_date in test_cases:
            with self.subTest(input=input_str):
                year, month, day = self.parser._get_ymd_time_by_str_save(input_str, base_date)
                result_date = date(year, month, day)
                self.assertEqual(result_date, expected_date)

    def test_edge_cases_week(self):
        """测试周的边界情况"""
        # 测试跨年、跨月的情况
        base_date = datetime(2023, 12, 31)  # 周日

        test_cases = [
            ("week +1 1", date(2024, 1, 1)),  # 下一周周一（新年）
            ("week -1 7", date(2023, 12, 24)),  # 上一周周日
        ]

        for input_str, expected_date in test_cases:
            with self.subTest(input=input_str):
                year, month, day = self.parser._get_ymd_time_by_str_save(input_str, base_date)
                result_date = date(year, month, day)
                self.assertEqual(result_date, expected_date)

    def test_invalid_today_format(self):
        """测试无效的today格式"""
        invalid_cases = [
            # "today ",  # 尾随空格
            # "td ",  # 尾随空格
            "today +",  # 不完整
            "td +",  # 不完整
            "today + abc",  # 非数字
            "td - def",  # 非数字
            "today+",  # 不完整无空格
            "td-",  # 不完整无空格
        ]

        for invalid_input in invalid_cases:
            with self.subTest(input=invalid_input):
                with self.assertRaises(ValueError):
                    self.parser._get_ymd_time_by_str_save(invalid_input)

    def test_invalid_week_format(self):
        """测试无效的week格式"""
        invalid_cases = [
            "week",  # 缺少周几
            "w",  # 缺少周几
            "week +",  # 不完整
            "w +",  # 不完整
            "week 8",  # 无效周几
            "w 8",  # 无效周几
            "week abc",  # 非数字周几
            "w + abc",  # 非数字周偏移
            "week+",  # 不完整无空格
            "w-",  # 不完整无空格
        ]

        for invalid_input in invalid_cases:
            with self.subTest(input=invalid_input):
                with self.assertRaises(ValueError):
                    self.parser._get_ymd_time_by_str_save(invalid_input)

    def test_case_insensitive(self):
        """测试大小写不敏感"""
        base_date = datetime(2023, 10, 15)

        test_cases = [
            ("TODAY", date(2023, 10, 15)),
            ("Today + 1", date(2023, 10, 16)),
            ("TD - 1", date(2023, 10, 14)),
            ("WEEK 1", date(2023, 10, 9)),
            ("Week +1 0", date(2023, 10, 22)),
        ]

        for input_str, expected_date in test_cases:
            with self.subTest(input=input_str):
                year, month, day = self.parser._get_ymd_time_by_str_save(input_str, base_date)
                result_date = date(year, month, day)
                self.assertEqual(result_date, expected_date)

    def test_whitespace_handling(self):
        """测试空格处理"""
        base_date = datetime(2023, 10, 15)

        test_cases = [
            ("  today  ", date(2023, 10, 15)),  # 前后空格
            ("  td + 1  ", date(2023, 10, 16)),  # 前后空格
            ("  week  1  ", date(2023, 10, 9)),  # 多空格
            ("  w  +1  0  ", date(2023, 10, 22)),  # 多空格
        ]

        for input_str, expected_date in test_cases:
            with self.subTest(input=input_str):
                year, month, day = self.parser._get_ymd_time_by_str_save(input_str, base_date)
                result_date = date(year, month, day)
                self.assertEqual(result_date, expected_date)

    def test_fallback_to_classic_parser(self):
        """测试回退到经典解析器"""
        # 这些应该由经典解析器处理
        test_cases = [
            "2020-2-29",  # 闰年
            "2020 02 29",  # 闰年
            "2021-2-28",  # 非闰年
        ]

        for input_str in test_cases:
            with self.subTest(input=input_str):
                # 应该能正常解析，不抛出异常
                try:
                    year, month, day = self.parser._get_ymd_time_by_str_save(input_str)
                    result_date = date(year, month, day)
                    self.assertIsInstance(result_date, date)
                except ValueError:
                    self.fail(f"经典解析器应该能处理: {input_str}")

    def test_invalid_dates_classic_parser(self):
        """测试经典解析器的无效日期"""
        invalid_cases = [
            "2020-13-1",  # 无效月份
            # "2020-2-30",  # 无效日期 29特殊判定
            # "2021-2-29",  # 非闰年2月29
            "abc-1-1",  # 非数字年
            "2020-abc-1",  # 非数字月
            "2020-1-abc",  # 非数字日
        ]

        for invalid_input in invalid_cases:
            with self.subTest(input=invalid_input):
                with self.assertRaises(ValueError):
                    self.parser._get_ymd_time_by_str_save(invalid_input)

    def test_empty_input(self):
        """测试空输入"""
        with self.assertRaises(ValueError):
            self.parser._get_ymd_time_by_str_save("")

    def test_none_input(self):
        """测试None输入"""
        with self.assertRaises(AttributeError):  # 因为会调用 strip()
            self.parser._get_ymd_time_by_str_save(None)


if __name__ == '__main__':
    unittest.main()