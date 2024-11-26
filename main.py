import json

import pandas as pd
import random
from datetime import datetime, timedelta
from openai import OpenAI

# 初始化OpenAI API客户端
client = OpenAI(api_key="", base_url='https://api.nextapi.fun')

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

# 随机生成一个时间范围（从五天前到当前时间）
def generate_random_time():
    time_diff = random.randint(0, 5 * 24 * 60 * 60)  # 5天以内的秒数
    random_time = datetime.now() - timedelta(seconds=time_diff)
    return random_time.strftime('%Y/%m/%d %H:%M:%S')

# 随机生成所用时间（以秒为单位）
def generate_random_duration():
    return f"{random.randint(60, 180)}秒"  # 生成 1-3 分钟之间的时间

# 随机生成来源、来源详情、IP
def generate_random_source():
    sources = ['手机提交', '电脑提交', '平板提交']
    source_details = ['直接访问', '通过链接访问', '社交媒体']
    ip_addresses = ['47.242.197.118(香港-未知)', '192.168.1.1(中国-北京)', '34.239.50.30(美国-弗吉尼亚州)']

    source = random.choice(sources)
    source_detail = random.choice(source_details)
    ip = random.choice(ip_addresses)

    return source, source_detail, ip

# 获取用户输入的生成数量
num_responses = int(input("请输入需要生成的问卷数量: "))

# 初始化存储问卷答案的列表
data = []

# 生成问卷并调用 OpenAI API 获取答案
for i in range(num_responses):
    # 构建问题文本
    questions_text = "\n".join(
        [f"{idx+1}、{q['question']}：{'；'.join(q['options'])}" for idx, q in enumerate(questions_with_options)]
    )

    # 调用 OpenAI API
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "system",
                "content": "您是一名问卷填写助手，请从给定的选项中选择答案。"
            },
            {
                "role": "user",
                "content": f"以下是问卷，请按要求选择答案：\n{questions_text}"
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
                            "description": "full answer item, include Option And text content, example: A.提高学历",
                            "items": {
                                "$ref": "#"
                            }
                        },
                    },
                    "required": ["answers_list"],
                }
            }
        }
    )

    # 解析答案
    answers = json.loads(response.choices[0].message.content)  # 假设这里是返回的答案 JSON
    answer_list = answers['answers_list']  # 获取答案列表

    # 获取随机数据
    submit_time = generate_random_time()
    duration = generate_random_duration()
    source, source_detail, ip = generate_random_source()

    # 组织回答格式
    answer_dict = {questions_with_options[i]["question"]: answer_list[i] for i in range(len(answer_list))}

    # 添加额外字段
    answer_dict['序号'] = i + 1
    answer_dict['提交答卷时间'] = submit_time
    answer_dict['所用时间'] = duration
    answer_dict['来源'] = source
    answer_dict['来源详情'] = source_detail
    answer_dict['来自IP'] = ip

    data.append(answer_dict)

    # 打印进度
    progress = (i + 1) / num_responses * 100
    print(f"生成进度: {progress:.2f}% ({i + 1}/{num_responses})")

# 将生成的数据转换为 DataFrame
df = pd.DataFrame(data)

# 输出到 Excel 文件
output_filename = f"问卷调查结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_filename, index=False)

print(f"所有数据生成完毕，已保存到 {output_filename}")
