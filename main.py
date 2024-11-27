import json
import pandas as pd
import random
from datetime import datetime, timedelta
from openai import OpenAI
from openpyxl import Workbook

# 初始化OpenAI API客户端
client = OpenAI(api_key="ak-", base_url='https://api.nextapi.fun')

# 定义问卷问题及选项
questions_with_options = [
    {"question": "您的性别", "options": ["A.男", "B.女"]},
    {"question": "您所在的年级", "options": ["A.大一", "B.大二", "C.大三", "D.大四"]},
    {"question": "您的专业", "options": ["A.经济管理类", "B.文史类", "C.理工类", "D.医学类", "E.艺术类", "F.信息技术计算机类"]},
    {"question": "大学毕业后是否有计划", "options": ["A.考研", "B.保研", "C.就业", "D.出国", "E.暂未考虑"]},
    {"question": "您是否做过职业规划", "options": ["A.无", "B.有，但比较粗糙", "C.有，比较清晰"]},
    {"question": "您对未来职业的期待是什么", "options": ["A.与所学专业有关", "B.薪资较高", "C.工作稳定", "D.发展空间大"]},
    {"question": "是否了解自己学校的保研条件", "options": ["A.了解", "B.部分了解", "C.不了解"]},
    {"question": "您选择考研或保研的原因是", "options": ["A.提高学历", "B.增加就业机会", "C.实现理想", "D.其他"]},
    {"question": "您选择直接就业的原因是", "options": ["A.就业机会好", "B.家庭经济需求", "C.其他"]},
    {"question": "选择就业的时候，您认为什么最重要", "options": ["A.薪水高低", "B.发展前景", "C.企业文化", "D.工作地点"]},
    {"question": "您认为当前社会对您的专业需求如何", "options": ["A.需求高", "B.需求一般", "C.需求低"]},
    {"question": "您通过哪些渠道了解职业信息和行业趋势", "options": ["A.网络媒体（如抖音、小红书）", "B.学校讲座", "C.职业规划老师", "D.实习/兼职经历"]},
    {"question": "您是否参与过与未来职业相关的实习、项目或竞赛", "options": ["A.是，多次参与", "B.是，少量参与", "C.未参与"]},
    {"question": "寝室风气如何影响您的学习态度和价值观", "options": ["A.积极影响", "B.消极影响", "C.一般，无明显积极或消极影响"]},
    {"question": "您对学校提供的生涯规划教育资源（如课程、讲座、咨询）的满意度如何", "options": ["A.非常满意", "B.满意", "C.一般", "D.不满意"]}
]

# 随机生成一个IP地址
def generate_random_ip():
    ip = f"{random.randint(1, 255)}.{random.randint(0, 255)}.{random.randint(0, 255)}.{random.randint(1, 255)}"
    guangxi_cities = [
        "(广西-南宁)", "(广西-桂林)", "(广西-柳州)", "(广西-梧州)", "(广西-北海)",
        "(广西-钦州)", "(广西-防城港)", "(广西-玉林)", "(广西-百色)", "(广西-贺州)",
        "(广西-河池)", "(广西-崇左)", "(广西-来宾)"
    ]
    location = random.choice(guangxi_cities)
    return f"{ip}{location}"

# 随机生成一个时间范围
def generate_random_time():
    time_diff = random.randint(0, 5 * 24 * 60 * 60)  # 5天以内的秒数
    random_time = datetime.now() - timedelta(seconds=time_diff)
    return random_time.strftime('%Y/%m/%d %H:%M:%S')

# 随机生成所用时间
def generate_random_duration():
    return f"{random.randint(60, 180)}秒"  # 生成 1-3 分钟之间的时间

# 随机生成来源信息
def generate_random_source():
    sources = ['手机提交']
    source_details = ['直接访问']
    source = random.choice(sources)
    source_detail = random.choice(source_details)
    ip = generate_random_ip()
    return source, source_detail, ip

# 获取用户输入的生成数量
num_responses = int(input("请输入需要生成的问卷数量: "))

# 初始化存储问卷答案的列表
data = []

# 创建 Excel 工作簿
output_filename = f"问卷调查结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
wb = Workbook()
ws = wb.active
headers = [q["question"] for q in questions_with_options] + ['序号', '提交答卷时间', '所用时间', '来源', '来源详情', '来自IP']
ws.append(headers)

# 生成问卷并调用 OpenAI API 获取答案
for i in range(num_responses):
    try:
        genders = ['男', '女']
        gender = random.choice(genders)
        grades = ['大一','大二','大三','大四']
        grade = random.choice(grades)
        questions_text = "\n".join(
            [f"{idx+1}、{q['question']}：{'；'.join(q['options'])}" for idx, q in enumerate(questions_with_options)]
        )

        # 调用 OpenAI API
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "我是一名问卷填写助手，可以从给定的选项中选择答案。"
                },
                {
                    "role": "user",
                    "content": f"你的性别: {gender}，您的年级: {grade}， 以下是问卷，请按要求选择答案：\n{questions_text}"
                }
            ],
            response_format={
                "type": "json_schema",
                "json_schema": {
                    "name": "answers",
                    "description": "answers",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "answers_list": {
                                "type": "array",
                                "description": "Array of indices representing selected options for each question (e.g., 0 for A, 1 for B).",
                                "items": {
                                    "type": "integer"
                                }
                            },
                        },
                        "required": ["answers_list"],
                    }
                }
            }
        )

        # 解析答案索引
        answers = json.loads(response.choices[0].message.content)
        answers_indices = answers['answers_list']  # 直接获取答案索引列表

        # 将索引转换为具体答案
        answers_text = [questions_with_options[q_idx]["options"][opt_idx] for q_idx, opt_idx in enumerate(answers_indices)]

        # 获取随机数据
        submit_time = generate_random_time()
        duration = generate_random_duration()
        source, source_detail, ip = generate_random_source()

        # 将具体答案写入 Excel
        row_data = answers_text + [i + 1, submit_time, duration, source, source_detail, ip]
        ws.append(row_data)
        wb.save(output_filename)

        # 打印进度
        progress = (i + 1) / num_responses * 100
        print(f"生成进度: {progress:.2f}% ({i + 1}/{num_responses})")

    except Exception as e:
        print(f"发生错误: {e}")

print(f"所有数据生成完毕，已保存到 {output_filename}")
