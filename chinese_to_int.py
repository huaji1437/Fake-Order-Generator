from typing import Optional

chinese_dict = {
    '零': 0, '〇': 0, '０': 0, "0": 0,
    '一': 1, '壹': 1, '１': 1, "1": 1,
    '二': 2, '两': 2, '贰': 2, '２': 2, "2": 2,
    '三': 3, '叁': 3, '３': 3, "3": 3,
    '四': 4, '肆': 4, '４': 4, "4": 4,
    '五': 5, '伍': 5, '５': 5, "5": 5,
    '六': 6, '陆': 6, '６': 6, "6": 6,
    '七': 7, '柒': 7, '７': 7, "7": 7,
    '八': 8, '捌': 8, '８': 8, "8": 8,
    '九': 9, '玖': 9, '９': 9, "9": 9,
}

num_to_chinese_arr = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']

all_chinese_num = ''.join(chinese_dict.keys())


def chinese_to_int(chinese_num) -> int:
    # 定义中文数字映射（包含大小写和全角半角）

    result = 0
    for char in chinese_num:
        if char in chinese_dict:
            result = result * 10 + chinese_dict[char]
        else:
            raise ValueError(f"无效的中文数字字符: {char}")

    return result

def chinese_to_int_op(chinese_num) -> Optional[int]:
    # 定义中文数字映射（包含大小写和全角半角）

    result = 0
    for char in chinese_num:
        if char in chinese_dict:
            result = result * 10 + chinese_dict[char]
        else:
            return None

    return result

def int_to_chinese_op(num: int) -> Optional[str]:
    if num < 0:
        raise ValueError("只能转换非负整数")
    if num == 0:
        return num_to_chinese_arr[0]
    chinese_num = ""
    while num > 0:
        digit = num % 10
        chinese_num = num_to_chinese_arr[digit] + chinese_num
        num //= 10
    return chinese_num