import io
import re
import sys
from abc import ABC, abstractmethod
from collections import defaultdict
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Any, Optional, Self, Iterable

import chinese_to_int
from chinese_to_int import chinese_to_int_op, int_to_chinese_op
from config_reader import ConfigReader
from docx_generator import DocumentGenerator

def print_red(text:str) -> None:
    print(f"\033[91m{text}\033[0m")


def print_warning(text:str) -> None:
    print(f"\033[93m⚠️ 警告: {text}\033[0m")



@contextmanager
def redirect_stdin_to_string(input_string: str):
    """将标准输入重定向到字符串的上下文管理器"""
    original_stdin = sys.stdin
    string_io = io.StringIO(input_string)
    sys.stdin = string_io
    try:
        yield
    finally:
        sys.stdin = original_stdin


def _parse_date_string(input_string: str) -> tuple[int, int, int]:
    """解析日期字符串，返回 (年, 月, 日)
    Raise:
        ValueError: 解析失败
    """

    # 模式1：严格模式，支持各种分隔符，允许1-2位数
    patterns = [
        r"(\d{4})[\D\s]*(\d{1,2})[\D\s]*(\d{1,2})",  # 2020-7-1, 2020 7 1, 2020/7/1
        r"(\d{4,})[\D\s]+(\d{1,2})[\D\s]*(\d{1,2})",  # 兜底：年份至少4位
    ]

    for pattern in patterns:
        match = re.match(pattern, input_string.strip())
        if match:
            year, month, day = map(int, match.groups())

            # 基本验证
            if 1 <= month <= 12 and 1 <= day <= 31:
                return year, month, day

    raise ValueError(f"无法解析日期: {input_string}")

class ABC_输入器(ABC):
    config_reader: ConfigReader
    docx_generator: Optional[DocumentGenerator]

    def for_mat_docx_and_pushout(self, *args, **kwargs) -> None:
        try:
            self.docx_generator.create_leave_form(*args, **kwargs)
        except Exception as e:
            print(f"生成请假单时发生错误: {e}")

    def __init__(self, config_reader: Optional[ConfigReader] = None):
        if config_reader is None:
            self.config_reader = ConfigReader("config.json")
        else:
            self.config_reader = config_reader
        self.docx_generator = DocumentGenerator(self.config_reader)

    def main(self) -> None:
        if self.docx_generator is None:
            raise ValueError("DocumentGenerator 未初始化")
        try:
            self._main()
        except Exception as e:
            print(f"main 发生错误: {e}")


    @abstractmethod
    def _main(self) -> int:
        return 0

    @abstractmethod
    def _get_test_input_head_string(self) -> str:
        """stu_data 前面的部分"""
        return ""

    def test_main(self) -> Self:
        with redirect_stdin_to_string(self._get_test_input_head_string()+self._get_接龙输入()):
            # self.docx_generator = None
            self._main()
        return self


    @classmethod
    def _get_接龙输入(cls) -> str:
        """主要是stu_data"""
        return """#接龙
视频组

1. 视频组24计应单2温正铁
2. 25数媒单2 梁思涵🪳
3. 25人工智能单 杨智睿
4. 25软件林则伽昊13736660120
5. 25软件魏宇剑
6. 25数媒  张心怡
7. 25数媒王玥
8. 25数媒陈怡然
9. 25计应 盛婕
10. 25数媒卢嘉铭
11. 25数媒奚玉镒   13868517461
12. 25数媒周弋松
13. 25网络单王墙
14. 25数媒单2刘楚鑫
15. 25数媒何籼增
16. 25软件钱宇程
17. 25软件技术胡书玮15724942093
18. 25软件李欣晨

"""

class 经典输入(ABC_输入器):

    def _get_test_input_head_string(self) -> str:
        return """2025.4.27
DH部
"""
    @classmethod
    def _get_接龙输入(cls) -> str:
        return """1. 视频组24计应单2温正铁
2. 25数媒单2 梁思涵🪳
3. 25人工智能单 杨智睿
4. 25软件林则伽昊13736660120
5. 25软件魏宇剑
6. 25数媒  张心怡
7. 25数媒王玥
8. 25数媒陈怡然
9. 25计应 盛婕
10. 25数媒卢嘉铭
11. 25数媒奚玉镒   13868517461
12. 25数媒周弋松
13. 25网络单王墙
14. 25数媒单2刘楚鑫
15. 25数媒何籼增
16. 25软件钱宇程
17. 25软件技术胡书玮15724942093
18. 25软件李欣晨

"""

    def _main(self) -> int:
        print("gitee：https://gitee.com/z_ky/Fake-orders.git")
        print("用于部门假单生成")
        print("数据来源：微信接龙复制/共享文档复制")
        print("注：若有错误请自行修改，上次更新时间：2025.11.13")
        input_time = input("请输入请假时间（例如：2025.4.27）：")
        cause = input("计信学院因xxx工作需要，以下同学需请假。(例：DH部)")
        input_time = input_time.split(".")
        year, month, day = int(input_time[0]), int(input_time[1]), int(input_time[2])
        stu_data:list[tuple[str,str]] = self.parse_student_data(self.get_student_input())

        self.for_mat_docx_and_pushout(stu_data, year=year, month=month, day=day, cause=cause)
        return 0

    @staticmethod
    def get_student_input():
        """获取用户输入的学生数据"""
        print("请输入学生数据（每行一个学生，格式如：23计应2xxx，输入空行结束）：")
        lines = []
        while True:
            line = input().strip()
            if not line:
                break
            lines.append(line)
        return "\n".join(lines)

    @staticmethod
    def parse_student_data(input_str: str):
        """使用正则表达式解析学生数据"""
        pattern = r'(?:^\d+\.\s*)?(\d{2})(云计算|计算机应用技术|计应|大数据技术|大数据|网络技术|网络|软件技术|软件|人工智能技术应用|人工智能|数字媒体技术班|数字媒体技术|数字媒体|数媒|电竞)(五年制)?(单)?(?:(\d+)(?:班|班级)?)?\s*?([\u4e00-\u9fa5]+)'
        matches = re.findall(pattern, input_str)
        students = []
        for match in matches:
            grade = match[0]  # 年级
            base_class = match[1]  # 基础专业名称
            is_five_year = match[2] if match[2] else ''  # 五年制标记
            is_single_class = match[3] if match[3] else ''  # 单班标记
            class_num = match[4] if match[4] else ''  # 阿拉伯数字班级号
            name = match[5].strip()  # 姓名
            modifiers = is_five_year + is_single_class
            class_mapping = {
                '网络技术': '网络',
                '计算机应用': '计应',
                '软件技术': '软件',
                '云计算': '云计算',
                '电子竞技': '电竞',
                '人工智能': '人工智能',
                '人工智能技术应用': '人工智能',
                '大数据技术': '大数据',
                '数字媒体': '数媒',
            }
            full_base_class = base_class
            for short, full in class_mapping.items():
                if short in base_class:
                    full_base_class = full
                    break
            class_name = full_base_class + modifiers + class_num
            full_class = f"{grade}{class_name}"
            students.append((full_class, name))
            print(f"{grade}{class_name} {name}")
        return students



class 我的输入器(ABC_输入器):

    # def test_日期输入器(self):
        # input_string_arr: dict[str, datetime.date] = {
        #     "asasc": False,
        #     "7asasii": False,
        #     "2020 2 1": True,
        #     "2020 2 2":,
        #     "202020 21 31",
        #     "2020 21 31",
        #     "today + 1",
        #     "td+1",
        #     "td-1",
        #     "week + 1 "
        # }

    def _get_ymd_time_by_str_save(self, input_string: str, base_date: Optional[datetime] = None
                                  ) -> tuple[int, int, int]:
        """
        解析灵活的日期字符串

        支持格式：
        - 标准日期: "2020-7-1", "2020/7/1", "2020 7 1"
        - 相对今天: "today + 1", "today", "td +1", "td -1"
        - 相对周: "week+1 0", "week-1 7" (0或7表示周日)

        Args:
            input_string: 输入的日期字符串
            base_date: 基准日期，默认为今天

        Returns:
            Tuple[int, int, int]: (年, 月, 日)

        Raises:
            ValueError: 当无法解析日期时
        """
        if base_date is None:
            base_date = datetime.now()

        # 清理输入字符串
        cleaned_input = input_string.strip().lower()

        # 1. 解析相对日期 (today/td [+- 天数]?)
        today_match = re.match(r'^(today|td)\s*(?:([+-])\s*(\d+))?$', cleaned_input)
        if today_match:
            days_offset = int(today_match.group(3)) if today_match.group(3) is not None else 0
            if today_match.group(2) == '-':
                days_offset = -days_offset
            target_date = base_date + timedelta(days=days_offset)
            return target_date.year, target_date.month, target_date.day
        if cleaned_input.startswith("today") or cleaned_input.startswith("td "):
            raise ValueError("today 格式如：(today/td [+- 天数]?)")
        # 2. 解析相对周 (week/w [+- 周数]? 星期几)
        # 不允许省略上周几，因为时间太长，容易忘记今天周几
        week_match = re.match(r'^(week|w)\s*(?:([+-])\s*(\d+))?\s*([0-7])$', cleaned_input)
        if week_match:
            weeks_offset = int(week_match.group(3)) if week_match.group(3) is not None else 0
            weekday = int(week_match.group(4))
            if week_match.group(2) == '-':
                weeks_offset = -weeks_offset


            # 将周日统一处理为7（Python中周一=0, 周日=6，我们调整为周日=7）
            if weekday == 0:
                weekday = 7
            weekday -= 1

            # 计算目标日期
            target_week_start = base_date - timedelta(days=base_date.weekday())
            target_date = target_week_start + timedelta(weeks=weeks_offset, days=weekday)
            return target_date.year, target_date.month, target_date.day
        if cleaned_input.startswith("week") or cleaned_input.startswith("w"):
            raise ValueError("week 格式如：(week/w [+- 周数]? 星期几)")

        return self._get_ymd_time_by_str_save_经典(input_string)

    def _get_ymd_time_by_str_save_经典(self, input_time_string: str
                                       ) -> tuple[int, int, int]:
        year: int = 0
        month: int = 0
        day: int = 0
        try:
            year, month, day = _parse_date_string(input_time_string)
        except ValueError as e:
            raise ValueError(f"日期解析错误: {e}") from e
        except Exception as e:
            raise ValueError(f"日期解析时遇到未知错误: {e}") from e
        else:
            return year, month, day

    pattern_one_cn_num = f"[\\d{chinese_to_int.all_chinese_num}]"
    pattern_sep_char = r"[\s,-;，-；_]"
    pattern = (
            r'^(?:\d+\.\s*)?([\u4e00-\u9fa5]+组)?(%s{2})\s*(?:(%s?)年制)?\s*([^,-;，-；\s班]+)\s*(?:(\d+)?(?:班|班级)?)?\s*[,-;，-；\s]\s*([\u4e00-\u9fa5]{1,8})'
            % (pattern_one_cn_num, pattern_one_cn_num)
    )
    @staticmethod
    def __to_tuple6_str(v:Any) -> tuple[str, str, str, str, str, str]:
        return v

    @classmethod
    def _match_line(cls, text: str) -> Optional[tuple[str, ...]]:
        match_result = re.match(cls.pattern, text)
        if not match_result:
            return None
        assert len(match_result.groups()) == 6
        result = cls.__to_tuple6_str(tuple((group if group is not None else "") for group in match_result.groups()))
        学年 = chinese_to_int_op(result[1]) or ""
        年制 = result[2] if not result[2].isdigit() else int_to_chinese_op(int(result[2])) or ""
        班级号 = result[4] if not result[4].isdigit() else int_to_chinese_op(int(result[4])) or ""
        return (result[0], str(学年), str(年制), result[3], str(班级号), result[5])

    def _iter_stu_data_match_result_from_input(self) -> Iterable[tuple[str, ...]]:
        """#接龙
视频组

1. 视频组24计应单2温正铁
2. 25数媒单2 梁思涵🪳
3. 25人工智能单 杨智睿
4. 25软件林则伽昊13736660120
5. 25软件魏宇剑
6. 25数媒  张心怡
7. 25数媒王玥
8. 25数媒陈怡然
9. 25计应 盛婕
10. 25数媒卢嘉铭
11. 25数媒奚玉镒   13868517461
12. 25数媒周弋松
13. 25网络单王墙
14. 25数媒单2刘楚鑫
15. 25数媒何籼增
16. 25软件钱宇程
17. 25软件技术胡书玮15724942093
18. 25软件李欣晨
"""
        print("请输入学生数据（每行一个学生，格式如：1. 23计应2xxx，输入空行结束）：")
        stu_data:list[tuple[str, str]] = []
        line:str = input().strip()
        while not line.startswith("1."):
            line = input().strip()
        while line:
            match_result = self._match_line(line)
            if not match_result:
                print_warning(f"未匹配的学生数据: {line}")
                line = input().strip()
                continue
            yield match_result
            line = input().strip()


    def _get_stu_data_from_input(self) -> list[tuple[str, str]]:
        stu_data:list[tuple[str, str]] = []
        for match_result in self._iter_stu_data_match_result_from_input():
            子分组, 学年, 年制, 专业名, 班级号, 姓名 = match_result
            专业名= self.config_reader.config["class_mappings"].get(专业名, 专业名)
            年制 = f"{年制}年制" if 年制 else ""
            # 班级号 = f"{班级号}" if 班级号 else ""
            班级名 = f'{专业名}{年制}{班级号}'
            完整班级名 = f"{学年}{班级名}"
            stu_data.append((完整班级名, 姓名))

        return stu_data

    def _get_test_input_head_string(self) -> str:
        return """202htehte2
2020 12 31
"""

    @classmethod
    def _get_接龙输入(cls) -> str:
        return """#接龙
我他妈乱七八糟在这里输入
都他妈没关系
这容错还能出错我直接气晕

空行插入测试

1. 视频组24计应单二-温正铁
3. 视频组二五计应单2 杨智睿
4. 视频组25软件2 林则伽昊13736660120
5. 软件组25软件,魏宇剑
6. 视频组25数媒;张心怡
7. 视频组25数媒 王玥
8. 软件组25数媒-陈怡然
9. 视频组25计应;盛婕
10. 视频组25数媒-卢嘉铭
11. 25数媒 奚玉镒   13868517461
12. 视频组25数媒 周弋松
13. 软件组25网络 单王墙
15. 软件组25数媒 何籼增
16. 软件组25软件二班 钱宇程
17. 软件组25软件技术 胡书玮15724942093
18. 25软件 李欣晨

"""

    def _get_stu_data_from_input_and_save_to_docx(self, *args,**kwargs) -> None:
        stu_data: list[tuple[str, str]] = self._get_stu_data_from_input()
        self.for_mat_docx_and_pushout(stu_data, *args,**kwargs)

    def _main(self) -> int:
            print("输入器开始")
            print("""请输入日期: %Y %m %d  如
    2020-7-1
    或 td - 1
    或  today + 1
    或week+1 0""")
            year: int = 0; month: int = 0; day: int = 0
            while True:
                try:
                    input_time_string = input()
                    # 若干空格 =
                    year, month, day = self._get_ymd_time_by_str_save(input_time_string)
                except ValueError as e:
                    print(f"日期解析错误: {e} 请重新输入")
                    continue
                except Exception as e:
                    print(f"日期解析时遇到未知错误: {e} 请重新输入")
                    continue
                else:
                    print(f"解析到的日期: {year}-{month}-{day}")


            cause : str = self.config_reader.get("cause", "？？部")
            self._get_stu_data_from_input_and_save_to_docx(year=year, month=month, day=day, cause=cause)
            return 0

class 分组多输出输入器(我的输入器):
    def _get_stu_data_from_input_and_save_to_docx(self, *args,cause = "", **kwargs) -> None:
        stu_data_grouped_by_子分组_dict: defaultdict[str, list[tuple[str, str]]] = defaultdict(list)
        for match_result in self._iter_stu_data_match_result_from_input():
            子分组, 学年, 年制, 专业名, 班级号, 姓名 = match_result
            专业名= self.config_reader.config["class_mappings"].get(专业名, 专业名)
            年制 = f"{年制}年制" if 年制 else ""
            # 班级号 = f"{班级号}" if 班级号 else ""
            班级名 = f'{专业名}{年制}{班级号}'
            完整班级名 = f"{学年}{班级名}"
            one_stu_data: tuple[str, str] = (完整班级名, 姓名)
            stu_data_grouped_by_子分组_dict[子分组].append(one_stu_data)

        for 子分组, stu_data in stu_data_grouped_by_子分组_dict.items():
            子分组 = 子分组 if 子分组 else "未分组"
            new_cause = f"{cause}{子分组}"
            self.for_mat_docx_and_pushout(stu_data, *args, cause=new_cause, **kwargs)


if __name__ == "__main__":
    # 经典输入().test_main()

    # my = 我的输入器()
    # my.config_reader.config["output_settings"]["file_name_format"] = "{year}年{month}月{day}日_{cause}假单_未定义输入器.docx"

    my = 分组多输出输入器()
    my.config_reader.config["output_settings"]["file_name_format"] = "{year}年{month}月{day}日_{cause}假单_分组多输出输入器.docx"

    my.test_main()
