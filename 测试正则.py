import re
import chinese_to_int


pattern_one_cn_num = f"[\\d{chinese_to_int.all_chinese_num}]"


pattern = (
        r'^(?:\d+\.\s*)?([\u4e00-\u9fa5]+组)?(%s{2})\s*(?:(%s?)年制)?\s*([^,-;，-；\s班]+)\s*(?:(\d+)?(?:班|班级)?)?\s*[,-;，-；\s]\s*([\u4e00-\u9fa5]{1,8})'
        % (pattern_one_cn_num, pattern_one_cn_num)
)
test_cases = [
    # "1. 23 四年制 计算机科学与技术2班-张三",
    # "2. 24 三年制 电子信息工程, 李四",
    # "3. 26 机械工程及自动化1班级; 手机号 --18的89",
    "4. 25 五年制 临床医学-你妈的232323-i9e2",
    "25 工商管理23班级-密码"
    "1. 24计应单2班-温正铁",
    "3. 视频组25人工智能单-杨智睿",
    "4. 软件组25软件班-林则伽昊13736660120",
    "5. 视频组25软件-魏宇剑",
    "6. 软件组25数媒-张心怡",
    "7. 垃圾25数媒-王玥",
    "8. 25数媒-陈怡然",
    # "9. 25计应-盛婕",
    # "10. 25数媒-卢嘉铭",
    # "11. 25数媒-奚玉镒   13868517461",
    # "12. 25数媒-周弋松",
    # "13. 25网络-单王墙",
    # "15. 25数媒-何籼增",
    # "16. 25软件-钱宇程",
    # "17. 25软件技术-胡书玮15724942093",
    # "18. 25软件-李欣晨",
]
def print_red(text):
    print(f"\033[91m{text}\033[0m")

def print_warning(text):
    print(f"\033[93m⚠️ 警告: {text}\033[0m")

for text in test_cases:
    match = re.match(pattern, text)
    if not match:
        print_red(f"文本: {text} 未匹配")
        continue
    print(f"文本: {text}")
    for i in range(1, 6):
        print(f"  组{i}: '{match.group(i)}'", end="\t")
    print("---")
    子组 = match.group(1)
    学年 = match.group(2)
    年制 = match.group(3)
    专业名 = match.group(4)
    班级号 = match.group(5)
    姓名 = match.group(6)
    assert len(match.groups()) == 6
    print(f"解析结果: {学年=}, {年制=}, {专业名=}, {班级号=}, {姓名=}")