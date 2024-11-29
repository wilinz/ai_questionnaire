import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import io
import os
from matplotlib import font_manager

# 读取问卷数据文件
file_path = "问卷调查结果_20241127_170952.xlsx"  # 请根据实际文件路径调整
df = pd.read_excel(file_path)

# 问卷问题列表
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

# 设置中文字体
font = font_manager.FontProperties(fname='/System/Library/Fonts/STHeiti Medium.ttc')  # 设置字体路径
plt.rcParams['font.family'] = font.get_name()  # 配置matplotlib字体

# 统计各个问题的选项分布并计算百分比
def count_options_percentage(df, column_name):
    counts = df[column_name].value_counts(normalize=True) * 100  # 计算百分比
    return counts

# 创建新的 Excel 文件
output_filename = f"问卷调查统计结果_{pd.to_datetime('today').strftime('%Y%m%d_%H%M%S')}.xlsx"
wb = Workbook()
ws = wb.active
ws.title = "统计数据"

# 写入统计数据的表头
headers = ['问题', '选项', '选项数量', '百分比']
ws.append(headers)

# 生成统计数据并写入 Excel
for question in questions_with_options:
    question_column = question['question']
    if question_column in df.columns:
        # 统计每个问题的选项百分比
        option_percentages = count_options_percentage(df, question_column)

        # 写入每个选项的统计信息
        for option, percentage in option_percentages.items():
            ws.append([question_column, option, option_percentages[option], f"{percentage:.2f}%"])

# 生成饼图并保存为图片
def generate_and_save_pie_chart(question_column, option_percentages, output_dir="charts"):
    plt.figure(figsize=(8, 6), dpi=300)  # 增加分辨率，尺寸设置为8x6英寸，300 DPI，适合插入Word
    plt.pie(option_percentages, labels=option_percentages.index, autopct='%1.2f%%', startangle=90, colors=sns.color_palette("Blues_d", len(option_percentages)))
    plt.title(f"{question_column} 选项分布", fontsize=16)
    plt.axis('equal')  # 使得饼图为圆形

    # 调整图表的边距
    plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

    # 确保图片保存的目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 生成以问题标题命名的图片文件路径
    image_path = os.path.join(output_dir, f"{question_column}.png")

    # 保存为图片
    plt.savefig(image_path, format='png')
    plt.close()

    return image_path

# 将图表插入到 Excel 文件
def insert_plot_to_excel(ws, image_path, row):
    img = Image(image_path)
    ws.add_image(img, f'E{row}')

# 在 Excel 文件的尾部插入图表
row = len(ws['A']) + 2  # 获取当前写入的行数，然后留出一行
for question in questions_with_options:
    question_column = question['question']
    if question_column in df.columns:
        option_percentages = count_options_percentage(df, question_column)
        image_path = generate_and_save_pie_chart(question_column, option_percentages)
        insert_plot_to_excel(ws, image_path, row)
        row += 40  # 增加图表之间的间隔，避免重叠

# 保存 Excel 文件
wb.save(output_filename)

print(f"所有统计数据和图表已保存到 {output_filename}")
