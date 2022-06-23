import re
import numpy as np
import pandas as pd
import xlrd
# from matplotlib import pyplot as plt
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from pandas import DataFrame
from xlrd import info
from PIL import Image
import xlsxwriter
import xlwt
import matplotlib.pyplot as plt
import os

walk = "D:/treasture/cloud_ama/report_automation/到家设计部MO分数toba-20220526.xlsx"
df = pd.read_excel(walk, engine='openpyxl', )  # dtype={'dept_code': str, 'parent_dept_code': str}
sheet_wenjuan = pd.read_excel(walk, sheet_name='1-问卷回收', header=0)  # sheet参数为None，返回的是全表
sheet_option = pd.read_excel(walk, sheet_name='2-选择题', header=0)  # sheet参数为None，返回的是全表
sheet_answer = pd.read_excel(walk, sheet_name='3-问答题', header=0)  # sheet参数为None，返回的是全表
# print(sheet_wenjuan)
#
# walk1 = "D:/treasture/cloud_ama/report_automation/报告正文1.xlsx"
# sheet_ZW1 = pd.read_excel(walk1,sheet_name=0,header=0,nrows=None)    # sheet参数为None，返回的是全表

#
# 一、问卷回收sheet
# 被评估人 D39
# 读取的字段名为”基本信息“，该列的第一个值为”姓名“
information_id = sheet_wenjuan['基本信息.2']
# print(information_id)
information_mis = sheet_wenjuan['基本信息.3']
information_name = sheet_wenjuan['基本信息.4']  # value_name[0]为第二个字段名   # <class 'pandas.core.series.Series'>
information_dept = sheet_wenjuan['基本信息.7']  # value_name[0]为第二个字段名   # <class 'pandas.core.series.Series'>
information_jp = sheet_wenjuan['基本信息.6']  # value_name[0]为第二个字段名   # <class 'pandas.core.series.Series'>
information_project = sheet_wenjuan['基本信息']  # value_name[0]为第二个字段名   # <class 'pandas.core.series.Series'>
information_time = sheet_wenjuan['基本信息.1']  # value_name[0]为第二个字段名   # <class 'pandas.core.series.Series'>

# 自评问卷数
questionnaire_self_send = sheet_wenjuan['自评问卷']
questionnaire_self_receive = sheet_wenjuan['自评问卷.1']
questionnaire_self_efficient = sheet_wenjuan['自评问卷.2']
# 上级问卷
questionnaire_superior_send = sheet_wenjuan['上级问卷']
questionnaire_superior_receive = sheet_wenjuan['上级问卷.1']
questionnaire_superior_efficient = sheet_wenjuan['上级问卷.2']
# 整体情况
questionnaire_overall_send = sheet_wenjuan['整体情况']
questionnaire_overall_receive = sheet_wenjuan['整体情况.1']
questionnaire_overall_efficient = sheet_wenjuan['整体情况.2']
questionnaire_overall_receiverate = sheet_wenjuan['整体情况.3']
questionnaire_overall_efficientrate = sheet_wenjuan['整体情况.4']
# 同级/合作伙伴问卷
questionnaire_companion_send = sheet_wenjuan['同级/合作伙伴问卷']
questionnaire_companion_receive = sheet_wenjuan['同级/合作伙伴问卷.1']
questionnaire_companion_efficient = sheet_wenjuan['同级/合作伙伴问卷.2']
# 下级问卷
questionnaire_lower_send = sheet_wenjuan['下级问卷']
questionnaire_lower_receive = sheet_wenjuan['下级问卷.1']
questionnaire_lower_efficient = sheet_wenjuan['下级问卷.2']

# 该列的第一个值为”姓名“
value_id = information_id[0]  # id
value_mis = information_mis[0]
value_name = information_name[0]  # 姓名
value_dept = information_dept[0]  # 评价时所在部门
value_jp = information_jp[0]  # Job Group
value_project = information_project[0]  # 项目名称
value_time = information_time[0]  # 报告日期
#

# 二、选择题sheet

# 上级均分-工作理念
work_phil_superior_0 = sheet_option['上级均分-工作理念']
work_phil_superior_1 = sheet_option['上级均分-工作理念.1']
work_phil_superior_2 = sheet_option['上级均分-工作理念.2']
work_phil_superior_3 = sheet_option['上级均分-工作理念.3']
work_phil_superior_4 = sheet_option['上级均分-工作理念.4']

# 上级均分-领导技能（19）
ls_skills_superior_0 = sheet_option['上级均分-领导技能']
ls_skills_superior_1 = sheet_option['上级均分-领导技能.1']
ls_skills_superior_2 = sheet_option['上级均分-领导技能.2']
ls_skills_superior_3 = sheet_option['上级均分-领导技能.3']
ls_skills_superior_4 = sheet_option['上级均分-领导技能.4']
ls_skills_superior_5 = sheet_option['上级均分-领导技能.5']
ls_skills_superior_6 = sheet_option['上级均分-领导技能.6']
ls_skills_superior_7 = sheet_option['上级均分-领导技能.7']
ls_skills_superior_8 = sheet_option['上级均分-领导技能.8']
ls_skills_superior_9 = sheet_option['上级均分-领导技能.9']
ls_skills_superior_10 = sheet_option['上级均分-领导技能.10']
ls_skills_superior_11 = sheet_option['上级均分-领导技能.11']
ls_skills_superior_12 = sheet_option['上级均分-领导技能.12']
ls_skills_superior_13 = sheet_option['上级均分-领导技能.13']
ls_skills_superior_14 = sheet_option['上级均分-领导技能.14']
ls_skills_superior_15 = sheet_option['上级均分-领导技能.15']
ls_skills_superior_16 = sheet_option['上级均分-领导技能.16']
ls_skills_superior_17 = sheet_option['上级均分-领导技能.17']
ls_skills_superior_18 = sheet_option['上级均分-领导技能.18']
ls_skills_superior_19 = sheet_option['上级均分-领导技能.19']

# 自评分数-领导技能（19）
ls_skills_self_0 = sheet_option['自评分数-领导技能']
ls_skills_self_1 = sheet_option['自评分数-领导技能.1']
ls_skills_self_2 = sheet_option['自评分数-领导技能.2']
ls_skills_self_3 = sheet_option['自评分数-领导技能.3']
ls_skills_self_4 = sheet_option['自评分数-领导技能.4']
ls_skills_self_5 = sheet_option['自评分数-领导技能.5']
ls_skills_self_6 = sheet_option['自评分数-领导技能.6']
ls_skills_self_7 = sheet_option['自评分数-领导技能.7']
ls_skills_self_8 = sheet_option['自评分数-领导技能.8']
ls_skills_self_9 = sheet_option['自评分数-领导技能.9']
ls_skills_self_10 = sheet_option['自评分数-领导技能.10']
ls_skills_self_11 = sheet_option['自评分数-领导技能.11']
ls_skills_self_12 = sheet_option['自评分数-领导技能.12']
ls_skills_self_13 = sheet_option['自评分数-领导技能.13']
ls_skills_self_14 = sheet_option['自评分数-领导技能.14']
ls_skills_self_15 = sheet_option['自评分数-领导技能.15']
ls_skills_self_16 = sheet_option['自评分数-领导技能.16']
ls_skills_self_17 = sheet_option['自评分数-领导技能.17']
ls_skills_self_18 = sheet_option['自评分数-领导技能.18']
ls_skills_self_19 = sheet_option['自评分数-领导技能.19']

# 同级/合作伙伴均分-领导技能
ls_skills_companion_0 = sheet_option['同级/合作伙伴均分-领导技能']
ls_skills_companion_1 = sheet_option['同级/合作伙伴均分-领导技能.1']
ls_skills_companion_2 = sheet_option['同级/合作伙伴均分-领导技能.2']
ls_skills_companion_3 = sheet_option['同级/合作伙伴均分-领导技能.3']
ls_skills_companion_4 = sheet_option['同级/合作伙伴均分-领导技能.4']
ls_skills_companion_5 = sheet_option['同级/合作伙伴均分-领导技能.5']
ls_skills_companion_6 = sheet_option['同级/合作伙伴均分-领导技能.6']
ls_skills_companion_7 = sheet_option['同级/合作伙伴均分-领导技能.7']
ls_skills_companion_8 = sheet_option['同级/合作伙伴均分-领导技能.8']
ls_skills_companion_9 = sheet_option['同级/合作伙伴均分-领导技能.9']
ls_skills_companion_10 = sheet_option['同级/合作伙伴均分-领导技能.10']
ls_skills_companion_11 = sheet_option['同级/合作伙伴均分-领导技能.11']
ls_skills_companion_12 = sheet_option['同级/合作伙伴均分-领导技能.12']
ls_skills_companion_13 = sheet_option['同级/合作伙伴均分-领导技能.13']
ls_skills_companion_14 = sheet_option['同级/合作伙伴均分-领导技能.14']
ls_skills_companion_15 = sheet_option['同级/合作伙伴均分-领导技能.15']
ls_skills_companion_16 = sheet_option['同级/合作伙伴均分-领导技能.16']
ls_skills_companion_17 = sheet_option['同级/合作伙伴均分-领导技能.17']
ls_skills_companion_18 = sheet_option['同级/合作伙伴均分-领导技能.18']
ls_skills_companion_19 = sheet_option['同级/合作伙伴均分-领导技能.19']

# 下级均分-领导技能（19）
ls_skills_lower_0 = sheet_option['下级均分-领导技能']
ls_skills_lower_1 = sheet_option['下级均分-领导技能.1']
ls_skills_lower_2 = sheet_option['下级均分-领导技能.2']
ls_skills_lower_3 = sheet_option['下级均分-领导技能.3']
ls_skills_lower_4 = sheet_option['下级均分-领导技能.4']
ls_skills_lower_5 = sheet_option['下级均分-领导技能.5']
ls_skills_lower_6 = sheet_option['下级均分-领导技能.6']
ls_skills_lower_7 = sheet_option['下级均分-领导技能.7']
ls_skills_lower_8 = sheet_option['下级均分-领导技能.8']
ls_skills_lower_9 = sheet_option['下级均分-领导技能.9']
ls_skills_lower_10 = sheet_option['下级均分-领导技能.10']
ls_skills_lower_11 = sheet_option['下级均分-领导技能.11']
ls_skills_lower_12 = sheet_option['下级均分-领导技能.12']
ls_skills_lower_13 = sheet_option['下级均分-领导技能.13']
ls_skills_lower_14 = sheet_option['下级均分-领导技能.14']
ls_skills_lower_15 = sheet_option['下级均分-领导技能.15']
ls_skills_lower_16 = sheet_option['下级均分-领导技能.16']
ls_skills_lower_17 = sheet_option['下级均分-领导技能.17']
ls_skills_lower_18 = sheet_option['下级均分-领导技能.18']
ls_skills_lower_19 = sheet_option['下级均分-领导技能.19']
#

# 他评均分-领导技能
ls_skills_other_0 = sheet_option['他评均分-领导技能']
ls_skills_other_1 = sheet_option['他评均分-领导技能.1']
ls_skills_other_2 = sheet_option['他评均分-领导技能.2']
ls_skills_other_3 = sheet_option['他评均分-领导技能.3']
ls_skills_other_4 = sheet_option['他评均分-领导技能.4']
ls_skills_other_5 = sheet_option['他评均分-领导技能.5']
ls_skills_other_6 = sheet_option['他评均分-领导技能.6']
ls_skills_other_7 = sheet_option['他评均分-领导技能.7']
ls_skills_other_8 = sheet_option['他评均分-领导技能.8']
ls_skills_other_9 = sheet_option['他评均分-领导技能.9']
ls_skills_other_10 = sheet_option['他评均分-领导技能.10']
ls_skills_other_11 = sheet_option['他评均分-领导技能.11']
ls_skills_other_12 = sheet_option['他评均分-领导技能.12']
ls_skills_other_13 = sheet_option['他评均分-领导技能.13']
ls_skills_other_14 = sheet_option['他评均分-领导技能.14']
ls_skills_other_15 = sheet_option['他评均分-领导技能.15']
ls_skills_other_16 = sheet_option['他评均分-领导技能.16']
ls_skills_other_17 = sheet_option['他评均分-领导技能.17']
ls_skills_other_18 = sheet_option['他评均分-领导技能.18']
ls_skills_other_19 = sheet_option['他评均分-领导技能.19']
#

# 第三部分问答题
mis_id = sheet_answer['mis号']
role = sheet_answer['角色']
advantage = sheet_answer['优点反馈']
disadvantage = sheet_answer['不足反馈']

# 报告中的固定文字

Introduction = '一、报告简介'  # B57+2
structure = '1、报告结构'
structure1 = '此报告是根据领导梯队对MO管理者的要求并综合了上级领导理念评估和360度领导技能评估结果之后形成的。'
structure2 = '本报告包含：①报告简介 ②测评结果综述 ③测评结果具体说明 ④发展建议 ⑤附录'
goal = '2、测评目的'
goal1 = '本次测评帮助你衡量你和领导梯队角色MO要求的一致性程度 ，有助于发现自己与领导梯队角色MO要求之间'
goal2 = '存在的成长机会，并帮助自身制定有针对性的发展计划以提升自己的管理能力。'
dimension = '3、测评维度'
dimension1 = '我们从2个方面去了解你在当前领导梯队角色MO上的匹配情况：'
dimension2 = '工作理念：管理者在工作中所践行的理念与价值取向'
dimension3 = '领导技能：领导梯队模型对MO的能力要求'
note = '4、测评结果得分说明'
answer = '问卷填答情况'

# 存储文件路径

file_walk = 'D:/treasture/cloud_ama/report_automation/mo_excel/'
# 文件
file_folder = 'D:/treasture/cloud_ama/report_automation/picture_report/'
# fulu
fulu = 'D:/treasture/cloud_ama/report_automation/'

if value_mis == 'mis号':
    for i in range(len(information_mis)):
        print('共计：' + str(len(information_mis) - 1) + '份报告')
        if i > 0:
            print('报告人：' + str(information_mis[i]))
            # print(type(ls_skills_companion_0[i]))
            # 创建文件夹、一个新Excel文件并添加一个工作表
            # file_folder = information_mis[i] +'/'
            #  if not os.path.isdir(file_folder):
            #      os.makedirs(file_folder)
            book_name = information_mis[i] + '.xlsx'
            # file_path = file_walk + file_folder + book_name
            file_path = file_walk + book_name
            book = xlsxwriter.Workbook(file_path)
            sheet_name = 'sheet' + str(i)
            sheet = book.add_worksheet(sheet_name)
            # print(sheet)
            # 报告标题样式
            property_name = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 18,  # 字号. 默认值 11
                'bold': True,  # 字体加粗
                'border': 0,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': False,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_name = book.add_format(property_name)
            # 报告内容样式
            property_content = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 8,  # 字号. 默认值 11
                'bold': False,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': False,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_content = book.add_format(property_content)
            property_content1 = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 7.5,  # 字号. 默认值 11
                'bold': True,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': False,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_content1 = book.add_format(property_content1)
            property_table = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 8,  # 字号. 默认值 11
                'bold': False,  # 字体加粗
                'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'vcenter',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_table = book.add_format(property_table)
            property_table_t = {
                'font_name': '微软雅黑',
                'font_size': 4.5,  # 字号. 默认值 11
                'bold': False,  # 字体加粗
                'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'vcenter',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_table_t = book.add_format(property_table_t)
            property_content_skill = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 6.5,  # 字号. 默认值 11
                'bold': False,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_content_skill = book.add_format(property_content_skill)

            property_content_skill_bold = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 9,  # 字号. 默认值 11
                'bold': True,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_content_skill_bold = book.add_format(property_content_skill_bold)
            property_title = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 8,  # 字号. 默认值 11
                'bold': True,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': False,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_title = book.add_format(property_title)
            property_title1 = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 8,  # 字号. 默认值 11
                'bold': True,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': False,  # 单元格内是否自动换行
                'color': 'black'
            }
            cell_format_title1 = book.add_format(property_title1)
            # 插入图片Demo
            sheet.insert_image('E6', '美团logo.png', {'x_scale': 0.5, 'y_scale': 0.5})  # {'x_scale': 0.5, 'y_scale': 0.5}
            # 插入标题
            report_title = "领导梯队（MO）360度评估报告"
            # 在向单元格中写入内容时，加上单元格样式
            sheet.write('C20', report_title, cell_format_name)
            # 报告人信息
            reporter = '被评估人： ' + str(information_name[i])
            department = '部门： ' + str(information_dept[i])
            jop_group = '序列： ' + str(information_jp[i])
            jop_project = '项目名称： ' + str(information_project[i])
            jop_time = '报告日期：' + str(information_time[i])
            sheet.write('D39', reporter, cell_format_content)
            sheet.write('D40', department, cell_format_content)
            sheet.write('D41', jop_group, cell_format_content)
            sheet.write('D42', jop_project, cell_format_content)
            sheet.write('D43', jop_time, cell_format_content)
            # 报告正文
            sheet.write('A56', Introduction, cell_format_title)
            sheet.write('B58', structure, cell_format_title)
            sheet.write('B59', structure1, cell_format_content)
            sheet.write('B60', structure2, cell_format_content)
            sheet.write('B62', goal, cell_format_title)
            sheet.write('B63', goal1, cell_format_content)
            sheet.write('B64', goal2, cell_format_content)
            sheet.write('B66', dimension, cell_format_title)
            sheet.write('B67', dimension1, cell_format_content)
            sheet.write('B68', dimension2, cell_format_content)
            sheet.write('B69', '领导技能：领导梯队模型对MO的能力要求', cell_format_content)
            # 4、测评结果得分说明
            sheet.write('B71', note, cell_format_title)
            sheet.merge_range('B73:C74', '    工作理念总分数区间', cell_format_table)
            sheet.merge_range('B75:C75', '         4≤分值≤5', cell_format_table)
            sheet.merge_range('B76:C76', '         3≤分值＜4', cell_format_table)
            sheet.merge_range('B77:C77', '         2≤分值＜3', cell_format_table)
            sheet.merge_range('B78:C78', '         1≤分值＜2', cell_format_table)

            sheet.merge_range('D73:E74', ' 与MO工作理念要求相比', cell_format_table)
            sheet.merge_range('D75:E75', '           非常一致', cell_format_table)
            sheet.merge_range('D76:E76', '           比较一致', cell_format_table)
            sheet.merge_range('D77:E77', '           有一些差距', cell_format_table)
            sheet.merge_range('D78:E78', '           有明显差距', cell_format_table)

            sheet.merge_range('F73:G74', '    领导技能总分数区间', cell_format_table)
            sheet.merge_range('F75:G75', '       3.88≤分值≤5', cell_format_table)
            sheet.merge_range('F76:G76', '      3.68≤分值<3.88', cell_format_table)
            sheet.merge_range('F77:G77', '      3.5≤分值<3.68', cell_format_table)
            sheet.merge_range('F78:G78', '          3.5≤分值', cell_format_table)

            sheet.merge_range('H73:I74', ' 与MO领导技能要求相比', cell_format_table)
            sheet.merge_range('H75:I75', '          有相对优势', cell_format_table)
            sheet.merge_range('H76:I76', '          符合期望', cell_format_table)
            sheet.merge_range('H77:I77', '          有一些差距', cell_format_table)
            sheet.merge_range('H78:I78', '          有明显差距', cell_format_table)

            # sheet.insert_image('B72', '工作理念&领导技能.png', {'x_scale': 1.1, 'y_scale': 1.1})
            sheet.write('B81', answer, cell_format_title)
            # 5、问卷填答情况
            sheet.merge_range('B83:C83', '      问卷情况/反馈人', cell_format_table)
            sheet.merge_range('B84:C84', '        问卷发放数', cell_format_table)
            sheet.merge_range('B85:C85', '        问卷回收数', cell_format_table)
            sheet.merge_range('B86:C86', '        无效问卷数', cell_format_table)
            sheet.write('D83', '     自己', cell_format_table)
            sheet.write('E83', '     上级', cell_format_table)
            sheet.merge_range('F83:G83', '         同级/合作伙伴', cell_format_table)
            sheet.write('H83', '     下级', cell_format_table)

            sheet.write('D84', questionnaire_self_send[i], cell_format_table)
            sheet.write('D85', questionnaire_self_receive[i], cell_format_table)
            # sheet.write('D86', questionnaire_self_efficient[i], cell_format_table)

            sheet.write('E84', questionnaire_superior_send[i], cell_format_table)
            sheet.write('E85', questionnaire_superior_receive[i], cell_format_table)
            # sheet.write('E86', questionnaire_superior_efficient[i], cell_format_table)

            sheet.merge_range('F84:G84', questionnaire_companion_send[i], cell_format_table)
            sheet.merge_range('F85:G85', questionnaire_companion_receive[i], cell_format_table)
            # sheet.merge_range('F86:G86', questionnaire_companion_efficient[i], cell_format_table)

            sheet.write('H84', questionnaire_lower_send[i], cell_format_table)
            sheet.write('H85', questionnaire_lower_receive[i], cell_format_table)
            # sheet.write('H86', questionnaire_lower_efficient[i], cell_format_table)

            sheet.write('D86', int(int(questionnaire_self_receive[i]) - int(questionnaire_self_efficient[i])),
                        cell_format_table)
            sheet.write('E86', int(int(questionnaire_superior_receive[i]) - int(questionnaire_superior_efficient[i])),
                        cell_format_table)
            sheet.merge_range('F86:G86',
                              int(int(questionnaire_companion_receive[i]) - int(questionnaire_companion_efficient[i])),
                              cell_format_table)
            sheet.write('H86', int(int(questionnaire_lower_receive[i]) - int(questionnaire_lower_efficient[i])),
                        cell_format_table)

            sheet.write('A89', '二、测评结果综述', cell_format_title)
            sheet.write('B91', '与领导梯队角色MO要求的比较', cell_format_title)
            sheet.write('D93', '工作理念', cell_format_title)
            # 画图
            #
            # plt 中文显示的问题
            plt.rcParams['font.sans-serif'] = ['SimHei']
            plt.rcParams['axes.unicode_minus'] = False
            #
            # 工作理念均分
            work_phil_superior_avg_1 = [work_phil_superior_0[i], work_phil_superior_1[i], work_phil_superior_2[i],
                                        work_phil_superior_3[i], work_phil_superior_4[i]]
            # print(type(work_phil_superior_1))
            print(work_phil_superior_avg_1)
            work_phil_superior_avg = []
            for wi in work_phil_superior_avg_1:
                if np.isnan(wi):
                    print('NULL')
                else:
                    work_phil_superior_avg.append(wi)

            work_phil_superior_avg = sum(work_phil_superior_avg) / len(work_phil_superior_avg)

            # 工作理念图片
            rects = plt.barh(work_phil_superior_avg, height=0.2, width=work_phil_superior_avg,
                             color='#FFC600')  # 横放条形图函数barh [orange]
            # fig = plt.figure(figsize=(5,5))
            # 加标注
            for rect in rects:
                width = rect.get_width()
                height = rect.get_height()
                # print(width)  # work_phil_superior_avg
                # print(height)  # 默认值0.8
                plt.text(width + 0.3, width, str(work_phil_superior_avg), ha='center', fontsize=18)
            #
            plt.xticks(range(0, 6, 1), fontsize=17)  # 刻度
            plt.yticks(range(0, 1, 1), color="black")
            # plt.title('工作理念')
            fig_str = str(information_name[i]) + '.工作理念.png'  # 更改图片路径
            plt.savefig(file_folder + '/' + fig_str, bbox_inches='tight')

            # plt.show()
            #
            sheet.insert_image('B94', file_folder + fig_str, {'x_scale': 0.4, 'y_scale': 0.4})
            plt.cla()
            # 工作理念总结
            if work_phil_superior_avg >= 4 and work_phil_superior_avg < 5:
                work_phil_score = '非常一致'
            elif work_phil_superior_avg >= 3 and work_phil_superior_avg < 4:
                work_phil_score = '比较一致'
            elif work_phil_superior_avg >= 2 and work_phil_superior_avg < 3:
                work_phil_score = '有一些差距'
            else:
                work_phil_score = '有明显差距'

            work_phil_text = '结果显示，你的工作理念与MO梯队要求' + work_phil_score
            sheet.write('F97', work_phil_text, cell_format_content)
            # 领导技能
            sheet.write('D108', '领导技能', cell_format_title)
            # 领导技能均分
            # 领导技能均分
            # ls_skills_other_avg = (float(ls_skills_other_0[i]) + float(ls_skills_other_1[i]) + float(
            #     ls_skills_other_2[i]) + float(ls_skills_other_3[i]) + float(ls_skills_other_4[i]) + float(
            #     ls_skills_other_5[i])
            #                        + float(ls_skills_other_6[i]) + float(ls_skills_other_7[i]) + float(
            #             ls_skills_other_8[i]) + float(ls_skills_other_9[i]) + float(ls_skills_other_10[i])
            #                        + float(ls_skills_other_11[i]) + float(ls_skills_other_12[i]) + float(
            #             ls_skills_other_13[i]) + float(ls_skills_other_14[i])
            #                        + float(ls_skills_other_15[i]) + float(ls_skills_other_16[i]) + float(
            #             ls_skills_other_17[i]) + float(ls_skills_other_18[i]) + float(ls_skills_other_19[i])) / 20
            # ls_skills_other_avg = round(ls_skills_other_avg, 2)
            ls_skills_other_avg = []
            ls_skills_other_avg_1 = [ls_skills_other_0[i], ls_skills_other_1[i], ls_skills_other_2[i],
                                     ls_skills_other_3[i], ls_skills_other_4[i]
                , ls_skills_other_5[i], ls_skills_other_6[i], ls_skills_other_7[i], ls_skills_other_8[i]
                , ls_skills_other_9[i], ls_skills_other_10[i], ls_skills_other_11[i], ls_skills_other_12[i]
                , ls_skills_other_13[i], ls_skills_other_14[i], ls_skills_other_15[i], ls_skills_other_16[i]
                , ls_skills_other_17[i], ls_skills_other_18[i], ls_skills_other_19[i]]
            for isoai in ls_skills_other_avg_1:
                if np.isnan(isoai):
                    print('NULL')
                else:
                    ls_skills_other_avg.append(isoai)

            ls_skills_other_avg = round(sum(ls_skills_other_avg) / len(ls_skills_other_avg), 2)

            rects_ls = plt.barh(ls_skills_other_avg, height=ls_skills_other_avg / 4, width=ls_skills_other_avg,
                                color='#FFC600', orientation="horizontal")  # 横放条形图函数 barh
            # 领导技能图片
            for rect_ls in rects_ls:
                width_ls = rect_ls.get_width()
                height_ls = rect_ls.get_height()
                plt.text(width_ls + 0.5, width_ls, str(ls_skills_other_avg), ha='center', fontsize=17)
            #
            plt.xticks(range(0, 6, 1), fontsize=17)  # 刻度
            plt.yticks(())
            # plt.title('工作理念')
            fig_str_ls = str(information_name[i]) + '.领导技能.png'  # 更改图片路径
            plt.savefig(file_folder + '/' + fig_str_ls, bbox_inches='tight')

            # plt.show()
            #
            sheet.insert_image('B109', file_folder + '/' + fig_str_ls, {'x_scale': 0.4, 'y_scale': 0.4})
            plt.cla()
            # 领导技能总结
            if ls_skills_other_avg >= 3.88:
                ls_skills_other_score = '有相对优势'
            elif ls_skills_other_avg >= 3.68 and ls_skills_other_avg < 3.88:
                ls_skills_other_score = '符合期望'
            elif ls_skills_other_avg >= 3.5 and ls_skills_other_avg < 3.68:
                ls_skills_other_score = '有一些差距'
            else:
                ls_skills_other_score = '有明显差距'

            ls_skills_other__text = '结果显示，你的领导技能与MO梯队要求相比' + ls_skills_other_score
            sheet.write('F112', ls_skills_other__text, cell_format_content)

            # 需要翻页
            # 文字部分
            sheet.write('B120', '1、工作理念', cell_format_title)
            sheet.write('B122', '评分规则回顾:', cell_format_content)
            sheet.write('B123', '1分 代表还没有，和上一个角色比没有任何变化，对新的理念缺乏认识或认同；', cell_format_content)
            sheet.write('B124', '2分 代表刚开始有，和上一个角色比有一些细微但积极的变化，开始认同新的理念；', cell_format_content)
            sheet.write('B125', '3分 代表一定程度有，和上一个角色比有比较明显且积极的变化，认同新的理念并在工作中有意识地运用；', cell_format_content)
            sheet.write('B126', '4分 代表明显有，和上一个角色比变化显著，高度认同新的理念，并在工作中能熟练运用；', cell_format_content)
            sheet.write('B127', '5分 代表知行合一，能够自知自觉地运用并正向影响周边的人（成为他们的模范）。', cell_format_content)
            sheet.write('B129', '本部分评估的是管理者的工作理念是否与MO的要求相吻合。领导梯队定义了各个层级管理者在工作理念上的差',
                        cell_format_content)
            sheet.write('B130', '异。你的上级对你的工作理念进行了评价。下面是TA的评价结果：',
                        cell_format_content)
            sheet.write('B131', 'MO的领导梯队角色要求：', cell_format_content)
            sheet.write('B132', '（1）真正认同管理工作的价值；', cell_format_content)
            sheet.write('B133', '（2）真正通过团队获得工作成果；', cell_format_content)
            sheet.write('B134', '（3）真正开始关注人，重视与人沟通和建立关系；', cell_format_content)
            sheet.write('B135', '（4）真正把注意力转到帮助他人和团队的成功上；', cell_format_content)
            sheet.write('B136', '（5）真正注重工作计划和强化团队执行力。', cell_format_content)

            # MO工作理念上级评估结果_图片
            m1 = '真正认同管理工作的价值'
            m2 = '真正通过团队获得工作成果'
            m3 = '真正开始关注人，重视与人沟通和建立关系'
            m4 = '真正把注意力转到帮助他人和团队的成功上'
            m5 = '真正注重工作计划和强化团队执行力'
            m1v = float(work_phil_superior_0[i])
            m2v = float(work_phil_superior_1[i])
            m3v = float(work_phil_superior_2[i])
            m4v = float(work_phil_superior_3[i])
            m5v = float(work_phil_superior_4[i])

            # 绘图数据准备
            mo = [m1, m2, m3, m4, m5]
            mo_value = [m1v, m2v, m3v, m4v, m5v]
            # 绘图
            # mo_value.reverse()

            b_mo = plt.barh(range(len(mo)), mo_value, color='#FFC600', tick_label=mo)
            # 添加数据标签
            for rect_mo in b_mo:
                width_ls = rect_mo.get_width()
                plt.text(width_ls + 0.1, rect_mo.get_y() + rect_mo.get_height() / 2, '%d' %
                         int(width_ls), ha='left', va='center')
            # 设置Y轴坐标轴上的刻度线标签
            # plt.yticks(range(len(mo)))
            # plt.ylabel(mo)
            plt.xticks(())
            fig_str_mo = str(information_name[i]) + 'MO工作理念上级评估结果.png'  # 更改图片路径
            plt.savefig(file_folder + '/' + fig_str_mo, bbox_inches='tight')
            # plt.show()
            #
            sheet.insert_image('B138', file_folder + '/' + fig_str_mo, {'x_scale': 0.75, 'y_scale': 0.75})
            plt.cla()
            # 测评结果
            sheet.write('B153', '测评结果显示你的工作理念和MO梯队要求一致的方面有：', cell_format_content)
            # 根据得分输出测评结果话术
            dict = {m1: m1v, m2: m2v, m3: m3v, m4: m4v, m5: m5v}
            mo_list1 = []
            loct1 = 154
            for k, v in dict.items():
                if v >= 4:
                    mo_list1.append(k)
                sheet.write_column('C154', mo_list1, cell_format_content1)
                len_mo_list1 = len(mo_list1)

            mo_list2 = []
            loct2 = loct1 + len_mo_list1 + 1
            for k, v in dict.items():
                if 3 <= v < 4:
                    mo_list2.append(k)
                sheet.write_column('C' + str(loct2 - 1), mo_list2, cell_format_content1)
                len_mo_list2 = len(mo_list2)
            if len_mo_list2 > 0:
                sheet.write('B' + str(loct2 - 1), "此外,", cell_format_content)
                sheet.write('B' + str(loct2 + len_mo_list2 - 1), "基本符合MO梯队要求。", cell_format_content)
            else:
                sheet.write('B' + str(loct2 + len_mo_list2), '因得分均小于3分，当前无与MO梯队要求一致的方面，后续需要加强观念的转变。',
                            cell_format_content)

            mo_list3 = []
            loct3 = loct2 + len_mo_list2 + 1
            for k, v in dict.items():
                if v < 3:
                    mo_list3.append(k)
                sheet.write_column('C' + str(loct3 + 1), mo_list3, cell_format_content1)
                len_mo_list3 = len(mo_list3)
            if len_mo_list3 > 0:
                sheet.write('B' + str(loct2 + len_mo_list2 + 1), "请注意，你还需要加强以下观念的转变：", cell_format_content)
            else:
                sheet.write('B' + str(loct2 + len_mo_list2 + 1), "你目前没有分数小于3分的工作理念", cell_format_content)

            # 领导技能
            sheet.write('B' + str(loct3 + len_mo_list2 + 3), "2、领导技能", cell_format_title)
            sheet.write('B' + str(loct3 + len_mo_list2 + 5), "评分规则回顾：", cell_format_content)
            sheet.write('B' + str(loct3 + len_mo_list2 + 6), "1分 代表存在严重短板，在该项上表现存在严重问题，迫切需要提高，否则会严重影响绩效;",
                        cell_format_content)
            sheet.write('B' + str(loct3 + len_mo_list2 + 7), "2分 代表有提升空间，在该项上表现存在一些问题，一定程度会影响绩效，需要予以关注；",
                        cell_format_content)
            sheet.write('B' + str(loct3 + len_mo_list2 + 8), "3分 代表基本符合期望，在该项上表现基本能够达到工作要求和公司的期望，无需特别关注；",
                        cell_format_content)
            sheet.write('B' + str(loct3 + len_mo_list2 + 9), "4分 代表有明显亮点，在该项上表现有明显的亮点，值得周围人学习；", cell_format_content)
            sheet.write('B' + str(loct3 + len_mo_list2 + 10), "5分 代表表现卓越，在该项上表现极为优秀，是被评估人标志性的优势，周围很少人能达到同等水平。",
                        cell_format_content)

            sheet.write('B' + str(loct3 + len_mo_list2 + 11),
                        "此部分评估的是管理者领导技能的表现是否与MO的要求相符。通过360度评估的题目收集他人对你的评价，",
                        cell_format_content)
            sheet.write('B' + str(loct3 + len_mo_list2 + 12),
                        "你可以直观看到自己的得分情况，更好地理解自身的相对优势与相对短板。",
                        cell_format_content)

            # 得分能力图例
            sheet.write('B' + str(loct3 + len_mo_list2 + 13), " 图例：", cell_format_content)
            sheet.insert_image('C' + str(loct3 + len_mo_list2 + 13), '得分能力图例.png', {'x_scale': 0.9, 'y_scale': 0.9})
            sheet.write('D' + str(loct3 + len_mo_list2 + 13), "代表各群体均分较低的能力", cell_format_content)
            sheet.write('D' + str(loct3 + len_mo_list2 + 14), "代表各群体均分较高的能力", cell_format_content)
            sheet.write('G' + str(loct3 + len_mo_list2 + 14), "注：被标记出的最低/最高能力不超过3项", cell_format_content)
            sheet.write('F' + str(loct3 + len_mo_list2 + 15), "领导技能得分明细表", cell_format_title)

            property_table1 = {
                # 'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 8,  # 字号. 默认值 11
                'bold': True,  # 字体加粗
                'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'vcenter',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black',
                'bg_color': '#cccccc'  # 颜色参数还需要改
            }
            cell_format_table1 = book.add_format(property_table1)  # 表头底色

            # 表格位置需要注意：整个表格需要连续
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 18)), "能力", cell_format_table1)
            print(str(loct3 + len_mo_list2 + 18))
            start_loc = str('B' + str(loct3 + len_mo_list2 + 18))
            end_loc = str('B' + str(loct3 + len_mo_list2 + 19))
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 19)), " ", cell_format_table1)
            sheet.merge_range('B178:B179', "   能力", cell_format_table1)
            # sheet.write(str('C' + str(loct3 + len_mo_list2 + 18)), "       行为描述", cell_format_table1)
            # sheet.write(str('C' + str(loct3 + len_mo_list2 + 19)), " ", cell_format_table1)
            sheet.merge_range('C178:E179', "           行为描述", cell_format_table1)
            # sheet.write(str('D' + str(loct3 + len_mo_list2 + 18)), "他评分数", cell_format_table1)
            # sheet.write(str('E' + str(loct3 + len_mo_list2 + 18)), " ", cell_format_table1)
            # sheet.write(str('F' + str(loct3 + len_mo_list2 + 18)), " ", cell_format_table1)
            # sheet.write(str('G' + str(loct3 + len_mo_list2 + 18)), " ", cell_format_table1)
            # sheet.write('G178', "他评分数", cell_format_table1)
            sheet.merge_range('F178:I178', "                  他评分数", cell_format_table1)
            sheet.write(str('F' + str(179)), "上级均分", cell_format_table1)
            sheet.write(str('G' + str(179)), "同级均分", cell_format_table1)
            sheet.write(str('H' + str(179)), "下级均分", cell_format_table1)
            sheet.write(str('I' + str(179)), "他评均分", cell_format_table1)
            # sheet.write(str('J' + str(loct3 + len_mo_list2 + 18)), "自评分数", cell_format_table1)
            # sheet.write(str('K' + str(loct3 + len_mo_list2 + 19)), "", cell_format_table1)
            sheet.merge_range('J178:J179', " 自评分数", cell_format_table1)

            # 表内内容
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 20)), "关注市场、客户和本领域动态", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 21)), " ", cell_format_table)
            sheet.merge_range('B180:B181', "关注市场、客户和本领域动态", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(180)) + ':' + str('E' + str(180)),
                "能定期收集并整合客户与市场的信息", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(181)) + ':' + str('E' + str(181)),
                "可以持续关注本专业&领域的最新动态，为团队提供输入", cell_format_table_t)
            # 纵向合并的同时也横向扩张
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 22)), "理解公司方向和部门策略重点", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 23)), " ", cell_format_table)
            sheet.merge_range('B182:B183', "理解公司方向和部门策略重点", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(182)) + ':' + str('E' + str(182)),
                "及时了解公司最新的发展方向与战略重点，并给团队分享", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(183)) + ':' + str('E' + str(183)),
                "能根据部门的策略重点有序的安排与推进工作", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 24)), "基于数据和事实分析、解决问题", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 25)), " ", cell_format_table)
            sheet.merge_range('B184:B185', "基于数据和事实分析、解决问题", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(184)) + ':' + str('E' + str(184)),
                "根据数据与事实有逻辑地分析问题，例如部门的目标设定标准与业务发展方向", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(185)) + ':' + str('E' + str(185)),
                "对部门的问题能提出专业有效的决策建议", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 26)), "识别高质量人才，帮助新人快速融入", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 27)), " ", cell_format_table)
            sheet.merge_range('B186:B187', "识别高质量人才，帮助新人快速融入", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(186)) + ':' + str('E' + str(186)),
                "善于观察与识别不同人的特质与长短板，甄别与吸引优秀的人才", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(187)) + ':' + str('E' + str(187)),
                "营造团队互信与友善的协作氛围，帮助新人快速融入团队", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 28)), "及时提供反馈、分享个人经验，帮助他人成长", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 29)), " ", cell_format_table)
            sheet.merge_range('B188:B189', "及时提供反馈、分享个人经验，帮助他人成长", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(188)) + ':' + str('E' + str(188)),
                "基于下属在工作中的行为给出及时与具体的激励型反馈与建设型反馈", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(189)) + ':' + str('E' + str(189)),
                "用合理的方式分享自身经验，辅导员工提升完成任务的能力", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 30)), "通过奖励和认可形成一个积极向上的团队氛围", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 31)), " ", cell_format_table)
            sheet.merge_range('B190:B191', "通过奖励和认可形成一个积极向上的团队氛围", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(190)) + ':' + str('E' + str(190)),
                "定期与下属沟通，并在过程中能持续帮助下属理解公司价值观", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(191)) + ':' + str('E' + str(191)),
                "能有效识别不同下属的激励点，并给出有针对性及有效的激励方式", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 32)), "通过有效分工、计划、跟进等确保高效执行", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 33)), " ", cell_format_table)
            sheet.merge_range('B192:B193', "通过有效分工、计划、跟进等确保高效执行", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(192)) + ':' + str('E' + str(192)),
                "根据团队成员的特点合理分配任务", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(193)) + ':' + str('E' + str(193)),
                "与下属共识详细任务实施计划，并及时跟踪检查进展情况", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 34)), "有行动力，能带领团队保质、保量、如期地完成任务", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 35)), " ", cell_format_table)
            sheet.merge_range('B194:B195', "有行动力，能带领团队保质、保量、如期地完成任务", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(194)) + ':' + str('E' + str(194)),
                "带领团队快速行动排除影响任务完成的障碍", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(195)) + ':' + str('E' + str(195)),
                "必要时及时调整计划和人员配置确保任务目标达成", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 36)), "切实通过复盘总结经验教训，不断改进", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 37)), " ", cell_format_table)
            sheet.merge_range('B196:B197', "切实通过复盘总结经验教训，不断改进", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(196)) + ':' + str('E' + str(196)),
                "引导团队通过规范的复盘方式不断总结经验与教训，养成复盘的习惯", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(197)) + ':' + str('E' + str(197)),
                "带领团队积极寻找持续优化工作方法与工作流程的机会", cell_format_table_t)

            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 38)), "通过主动找不足和差距不断提升自己", cell_format_table)
            # sheet.write(str('B' + str(loct3 + len_mo_list2 + 39)), " ", cell_format_table)
            sheet.merge_range('B198:B199', "通过主动找不足和差距不断提升自己", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(198)) + ':' + str('E' + str(198)),
                "主动寻求并接纳他人的意见和反馈，不断自我改进", cell_format_table_t)
            sheet.merge_range(
                str('C' + str(199)) + ':' + str('E' + str(199)),
                "可以持续探索和学习新的知识和领域", cell_format_table_t)
            #

            # 填数（上级评分）
            ls_skills_superior_avg_01 = []
            if np.isnan(ls_skills_superior_0[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_01.append(ls_skills_superior_0[i])

            if np.isnan(ls_skills_superior_1[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_01.append(ls_skills_superior_1[i])

            if len(ls_skills_superior_avg_01) == 0:
                ls_skills_superior_avg_01 = 0
            else:
                ls_skills_superior_avg_01 = round(
                    sum(ls_skills_superior_avg_01) / len(ls_skills_superior_avg_01), 2)
            # ls_skills_superior_avg_01 = (ls_skills_superior_2[i] + ls_skills_superior_1[i]) / 2
            # ls_skills_superior_avg_01 = round(ls_skills_superior_avg_01, 2)
            ls_skills_superior_avg_23 = []
            if np.isnan(ls_skills_superior_2[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_23.append(ls_skills_superior_2[i])

            if np.isnan(ls_skills_superior_3[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_23.append(ls_skills_superior_3[i])

            if len(ls_skills_superior_avg_23) == 0:
                ls_skills_superior_avg_23 = 0
            else:
                ls_skills_superior_avg_23 = round(
                    sum(ls_skills_superior_avg_23) / len(ls_skills_superior_avg_23), 2)
            # ls_skills_superior_avg_23 = (ls_skills_superior_2[i] + ls_skills_superior_3[i]) / 2
            # ls_skills_superior_avg_23 = round(ls_skills_superior_avg_23, 2)
            # sheet.merge_range('D182:D183', ls_skills_superior_avg_23, cell_format_table)
            # sheet.write(str('D' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_23, cell_format_content)
            ls_skills_superior_avg_45 = []
            if np.isnan(ls_skills_superior_4[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_45.append(ls_skills_superior_4[i])

            if np.isnan(ls_skills_superior_5[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_45.append(ls_skills_superior_5[i])

            if len(ls_skills_superior_avg_45) == 0:
                ls_skills_superior_avg_45 = 0
            else:
                ls_skills_superior_avg_45 = round(
                    sum(ls_skills_superior_avg_45) / len(ls_skills_superior_avg_45), 2)
            # ls_skills_superior_avg_45 = (ls_skills_superior_4[i] + ls_skills_superior_5[i]) / 2
            # ls_skills_superior_avg_45 = round(ls_skills_superior_avg_45, 2)
            # sheet.merge_range('D184:D185', ls_skills_superior_avg_45, cell_format_table)
            # sheet.write(str('E' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_23, cell_format_content)
            ls_skills_superior_avg_67 = []
            if np.isnan(ls_skills_superior_6[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_67.append(ls_skills_superior_6[i])

            if np.isnan(ls_skills_superior_7[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_67.append(ls_skills_superior_7[i])

            if len(ls_skills_superior_avg_67) == 0:
                ls_skills_superior_avg_67 = 0
            else:
                ls_skills_superior_avg_67 = round(
                    sum(ls_skills_superior_avg_67) / len(ls_skills_superior_avg_67), 2)
            # ls_skills_superior_avg_67 = (ls_skills_superior_6[i] + ls_skills_superior_7[i]) / 2
            # ls_skills_superior_avg_67 = round(ls_skills_superior_avg_67, 2)
            # sheet.merge_range('D186:D187', ls_skills_superior_avg_67, cell_format_table)
            ls_skills_superior_avg_89 = []
            if np.isnan(ls_skills_superior_8[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_89.append(ls_skills_superior_8[i])

            if np.isnan(ls_skills_superior_9[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_89.append(ls_skills_superior_9[i])

            if len(ls_skills_superior_avg_89) == 0:
                ls_skills_superior_avg_89 = 0
            else:
                ls_skills_superior_avg_89 = round(
                    sum(ls_skills_superior_avg_89) / len(ls_skills_superior_avg_89), 2)
            # ls_skills_superior_avg_89 = (ls_skills_superior_8[i] + ls_skills_superior_9[i]) / 2
            # ls_skills_superior_avg_89 = round(ls_skills_superior_avg_89, 2)
            # sheet.merge_range('D188:D189', ls_skills_superior_avg_89, cell_format_table)
            ls_skills_superior_avg_1011 = []
            if np.isnan(ls_skills_superior_10[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1011.append(ls_skills_superior_10[i])

            if np.isnan(ls_skills_superior_11[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1011.append(ls_skills_superior_11[i])

            if len(ls_skills_superior_avg_1011) == 0:
                ls_skills_superior_avg_1011 = 0
            else:
                ls_skills_superior_avg_1011 = round(
                    sum(ls_skills_superior_avg_1011) / len(ls_skills_superior_avg_1011), 2)
            # ls_skills_superior_avg_1011 = (ls_skills_superior_10[i] + ls_skills_superior_11[i]) / 2
            # ls_skills_superior_avg_1011 = round(ls_skills_superior_avg_1011, 2)
            # sheet.merge_range('D190:D191', ls_skills_superior_avg_1011, cell_format_table)
            ls_skills_superior_avg_1213 = []
            if np.isnan(ls_skills_superior_12[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1213.append(ls_skills_superior_12[i])

            if np.isnan(ls_skills_superior_13[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1213.append(ls_skills_superior_13[i])

            if len(ls_skills_superior_avg_1213) == 0:
                ls_skills_superior_avg_1213 = 0
            else:
                ls_skills_superior_avg_1213 = round(
                    sum(ls_skills_superior_avg_1213) / len(ls_skills_superior_avg_1213), 2)
            # ls_skills_superior_avg_1213 = (ls_skills_superior_12[i] + ls_skills_superior_13[i]) / 2
            # ls_skills_superior_avg_1213 = round(ls_skills_superior_avg_1213, 2)
            # sheet.merge_range('D192:D193', ls_skills_superior_avg_1213, cell_format_table)
            ls_skills_superior_avg_1415 = []
            if np.isnan(ls_skills_superior_14[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1415.append(ls_skills_superior_14[i])

            if np.isnan(ls_skills_superior_15[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1415.append(ls_skills_superior_15[i])

            if len(ls_skills_superior_avg_1415) == 0:
                ls_skills_superior_avg_1415 = 0
            else:
                ls_skills_superior_avg_1415 = round(
                    sum(ls_skills_superior_avg_1415) / len(ls_skills_superior_avg_1415), 2)
            # ls_skills_superior_avg_1415 = (ls_skills_superior_14[i] + ls_skills_superior_15[i]) / 2
            # ls_skills_superior_avg_1415 = round(ls_skills_superior_avg_1415, 2)
            # sheet.merge_range('D194:D195', ls_skills_superior_avg_1415, cell_format_table)
            ls_skills_superior_avg_1617 = []
            if np.isnan(ls_skills_superior_16[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1617.append(ls_skills_superior_16[i])

            if np.isnan(ls_skills_superior_17[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1617.append(ls_skills_superior_17[i])

            if len(ls_skills_superior_avg_1617) == 0:
                ls_skills_superior_avg_1617 = 0
            else:
                ls_skills_superior_avg_1617 = round(
                    sum(ls_skills_superior_avg_1617) / len(ls_skills_superior_avg_1617), 2)
            # ls_skills_superior_avg_1617 = (ls_skills_superior_16[i] + ls_skills_superior_17[i]) / 2
            # ls_skills_superior_avg_1617 = round(ls_skills_superior_avg_1617, 2)
            # sheet.merge_range('D196:D197', ls_skills_superior_avg_1617, cell_format_table)
            ls_skills_superior_avg_1819 = []
            if np.isnan(ls_skills_superior_18[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1819.append(ls_skills_superior_18[i])

            if np.isnan(ls_skills_superior_19[i]):
                print('NAN')
            else:
                ls_skills_superior_avg_1819.append(ls_skills_superior_19[i])

            if len(ls_skills_superior_avg_1819) == 0:
                ls_skills_superior_avg_1819 = 0
            else:
                ls_skills_superior_avg_1819 = round(
                    sum(ls_skills_superior_avg_1819) / len(ls_skills_superior_avg_1819), 2)
            # ls_skills_superior_avg_1819 = (ls_skills_superior_18[i] + ls_skills_superior_19[i]) / 2
            # ls_skills_superior_avg_1819 = round(ls_skills_superior_avg_1819, 2)
            # sheet.merge_range('D198:D199', ls_skills_superior_avg_1819, cell_format_table)
            # 表格填充底色
            # 深橙色
            property_table_color_deep = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 7.5,  # 字号. 默认值 11
                # 'bold': True,  # 字体加粗
                'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'vcenter',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                # 'color': 'black',
                'fg_color': '#FFC600',  # 颜色参数还需要改
            }
            # 浅橙色
            property_table_color_shallow = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 7.5,  # 字号. 默认值 11
                # 'bold': True,  # 字体加粗
                'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'vcenter',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black',
                'fg_color': '#fbeeac'  # 颜色参数: https://www.cnblogs.com/darkknightzh/p/6117528.html
            }
            # 无色
            property_table_color_blank = {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 7.5,  # 字号. 默认值 11
                # 'bold': True,  # 字体加粗
                'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'vcenter',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                'color': 'black'
            }

            # 排序判断(上级评分)
            ls_skills_superior_list_1 = [ls_skills_superior_avg_01, ls_skills_superior_avg_23,
                                         ls_skills_superior_avg_45, ls_skills_superior_avg_67,
                                         ls_skills_superior_avg_89, ls_skills_superior_avg_1011,
                                         ls_skills_superior_avg_1213, ls_skills_superior_avg_1415,
                                         ls_skills_superior_avg_1617, ls_skills_superior_avg_1819]
            ls_skills_superior_list = []
            for si in ls_skills_superior_list_1:
                if si == 0:
                    print('NULL')
                else:
                    ls_skills_superior_list.append(si)
            # print(1)
            print('上级评分：' + str(ls_skills_superior_list))

            #
            format1 = []
            # 从小到大排序
            dic1_sort = np.argsort(ls_skills_superior_list)  # 用索引引用字典  dict[i]
            dss_len = len(ls_skills_superior_list)
            # 索引排列
            # 取索引的后四位（最大的4位）
            i4 = dic1_sort[dss_len - 4]  # 第4位
            i3 = dic1_sort[dss_len - 3]  # 第3位
            i2 = dic1_sort[dss_len - 2]  # 第2位
            i1 = dic1_sort[dss_len - 1]  # 第1位
            # 取索引的前四位（最小的4位）
            i7 = dic1_sort[3]  # 第4位
            i8 = dic1_sort[2]  # 第3位
            i9 = dic1_sort[1]  # 第2位
            i10 = dic1_sort[0]  # 第1位
            #

            # 判断前3位和后3位的逻辑分开写
            # 先预置单元格格式为不添加任何颜色
            D1 = book.add_format(property_table_color_blank)
            D2 = book.add_format(property_table_color_blank)
            D3 = book.add_format(property_table_color_blank)
            D4 = book.add_format(property_table_color_blank)
            D5 = book.add_format(property_table_color_blank)
            D6 = book.add_format(property_table_color_blank)
            D7 = book.add_format(property_table_color_blank)
            D8 = book.add_format(property_table_color_blank)
            D9 = book.add_format(property_table_color_blank)
            D10 = book.add_format(property_table_color_blank)
            #

            # 判断位数(前3位）
            if (ls_skills_superior_list[i1] == ls_skills_superior_list[i2] and ls_skills_superior_list[i2] ==
                ls_skills_superior_list[i3] and ls_skills_superior_list[i3] != ls_skills_superior_list[i4]) or \
                    (ls_skills_superior_list[i1] != ls_skills_superior_list[i2] and ls_skills_superior_list[i2] ==
                     ls_skills_superior_list[i3] and ls_skills_superior_list[i3] != ls_skills_superior_list[i4]) or \
                    (ls_skills_superior_list[i1] == ls_skills_superior_list[i2] and ls_skills_superior_list[i2] !=
                     ls_skills_superior_list[i3] and ls_skills_superior_list[i3] != ls_skills_superior_list[i4]) \
                    or (ls_skills_superior_list[i1] != ls_skills_superior_list[i2] and ls_skills_superior_list[i2] !=
                        ls_skills_superior_list[i3] and ls_skills_superior_list[i3] != ls_skills_superior_list[i4]):
                # 有最大的前3项,定位到最大单元格，然后标出颜色
                # 确定最大值，如何关联到颜色（考虑用键值对的形式）例如找到最大值x，那么其对应的y可以被赋值
                if ls_skills_superior_avg_01 == ls_skills_superior_list[i1] or ls_skills_superior_avg_01 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_01 == ls_skills_superior_list[i3]:
                    D1 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_23 == ls_skills_superior_list[i1] or ls_skills_superior_avg_23 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_23 == ls_skills_superior_list[i3]:
                    D2 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_45 == ls_skills_superior_list[i1] or ls_skills_superior_avg_45 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_45 == ls_skills_superior_list[i3]:
                    D3 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_67 == ls_skills_superior_list[i1] or ls_skills_superior_avg_67 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_67 == ls_skills_superior_list[i3]:
                    D4 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_89 == ls_skills_superior_list[i1] or ls_skills_superior_avg_89 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_89 == ls_skills_superior_list[i3]:
                    D5 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1011 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1011 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_1011 == ls_skills_superior_list[i3]:
                    D6 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1213 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1213 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_1213 == ls_skills_superior_list[i3]:
                    D7 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1415 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1415 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_1415 == ls_skills_superior_list[i3]:
                    D8 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1617 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1617 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_1617 == ls_skills_superior_list[i3]:
                    D9 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1819 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1819 == \
                        ls_skills_superior_list[i2] or ls_skills_superior_avg_1819 == ls_skills_superior_list[i3]:
                    D10 = book.add_format(property_table_color_deep)

            # 最大值有两个（考虑的情况没有穷尽）
            if (ls_skills_superior_list[i1] == ls_skills_superior_list[i2] and ls_skills_superior_list[i3] == \
                ls_skills_superior_list[i4]) or (ls_skills_superior_list[i1] != ls_skills_superior_list[i2] and
                                                 ls_skills_superior_list[i2] != ls_skills_superior_list[i3] and
                                                 ls_skills_superior_list[i3] == ls_skills_superior_list[i4]):
                if ls_skills_superior_avg_01 == ls_skills_superior_list[i1] or ls_skills_superior_avg_01 == \
                        ls_skills_superior_list[i2]:
                    D1 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_23 == ls_skills_superior_list[i1] or ls_skills_superior_avg_23 == \
                        ls_skills_superior_list[i2]:
                    D2 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_45 == ls_skills_superior_list[i1] or ls_skills_superior_avg_45 == \
                        ls_skills_superior_list[i2]:
                    D3 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_67 == ls_skills_superior_list[i1] or ls_skills_superior_avg_67 == \
                        ls_skills_superior_list[i2]:
                    D4 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_89 == ls_skills_superior_list[i1] or ls_skills_superior_avg_89 == \
                        ls_skills_superior_list[i2]:
                    D5 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1011 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1011 == \
                        ls_skills_superior_list[i2]:
                    D6 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1213 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1213 == \
                        ls_skills_superior_list[i2]:
                    D7 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1415 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1415 == \
                        ls_skills_superior_list[i2]:
                    D8 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1617 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1617 == \
                        ls_skills_superior_list[i2]:
                    D9 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1819 == ls_skills_superior_list[i1] or ls_skills_superior_avg_1819 == \
                        ls_skills_superior_list[i2]:
                    D10 = book.add_format(property_table_color_deep)

            # 最大值只有一个
            if (ls_skills_superior_list[i1] != ls_skills_superior_list[i2] and
                    ls_skills_superior_list[i2] == ls_skills_superior_list[i3] and
                    ls_skills_superior_list[i2] == ls_skills_superior_list[i4]):
                if ls_skills_superior_avg_01 == ls_skills_superior_list[i1]:
                    D1 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_23 == ls_skills_superior_list[i1]:
                    D2 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_45 == ls_skills_superior_list[i1]:
                    D3 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_67 == ls_skills_superior_list[i1]:
                    D4 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_89 == ls_skills_superior_list[i1]:
                    D5 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1011 == ls_skills_superior_list[i1]:
                    D6 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1213 == ls_skills_superior_list[i1]:
                    D7 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1415 == ls_skills_superior_list[i1]:
                    D8 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1617 == ls_skills_superior_list[i1]:
                    D9 = book.add_format(property_table_color_deep)
                if ls_skills_superior_avg_1819 == ls_skills_superior_list[i1]:
                    D10 = book.add_format(property_table_color_deep)

            # 无最大
            if (ls_skills_superior_list[i1] == ls_skills_superior_list[i2] and
                    ls_skills_superior_list[i2] == ls_skills_superior_list[i3] and
                    ls_skills_superior_list[i3] == ls_skills_superior_list[i4]):
                D1 = book.add_format(property_table_color_blank)
                D2 = book.add_format(property_table_color_blank)
                D3 = book.add_format(property_table_color_blank)
                D4 = book.add_format(property_table_color_blank)
                D5 = book.add_format(property_table_color_blank)
                D6 = book.add_format(property_table_color_blank)
                D7 = book.add_format(property_table_color_blank)
                D8 = book.add_format(property_table_color_blank)
                D9 = book.add_format(property_table_color_blank)
                D10 = book.add_format(property_table_color_blank)

            # 后三位判断
            if (ls_skills_superior_list[i10] == ls_skills_superior_list[i9] and ls_skills_superior_list[i9] ==
                ls_skills_superior_list[i8] and ls_skills_superior_list[i8] != ls_skills_superior_list[i7]) or \
                    (ls_skills_superior_list[i10] != ls_skills_superior_list[i9] and ls_skills_superior_list[i9] ==
                     ls_skills_superior_list[i8] and
                     ls_skills_superior_list[i8] != ls_skills_superior_list[i7]) or \
                    (ls_skills_superior_list[i10] == ls_skills_superior_list[i9] and ls_skills_superior_list[i9] !=
                     ls_skills_superior_list[i8] and ls_skills_superior_list[i8] != ls_skills_superior_list[i7]) or \
                    (ls_skills_superior_list[i10] != ls_skills_superior_list[i9] and ls_skills_superior_list[i9] !=
                     ls_skills_superior_list[i8] and ls_skills_superior_list[i8] != ls_skills_superior_list[i7]):

                if ls_skills_superior_avg_01 == ls_skills_superior_list[i10] or ls_skills_superior_avg_01 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_01 == ls_skills_superior_list[i8]:
                    D1 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_23 == ls_skills_superior_list[i10] or ls_skills_superior_avg_23 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_23 == ls_skills_superior_list[i8]:
                    D2 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_45 == ls_skills_superior_list[i10] or ls_skills_superior_avg_45 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_45 == ls_skills_superior_list[i8]:
                    D3 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_67 == ls_skills_superior_list[i10] or ls_skills_superior_avg_67 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_67 == ls_skills_superior_list[i8]:
                    D4 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_89 == ls_skills_superior_list[i10] or ls_skills_superior_avg_89 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_89 == ls_skills_superior_list[i8]:
                    D5 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1011 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1011 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_1011 == ls_skills_superior_list[i8]:
                    D6 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1213 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1213 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_1213 == ls_skills_superior_list[i8]:
                    D7 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1415 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1415 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_1415 == ls_skills_superior_list[i8]:
                    D8 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1617 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1617 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_1617 == ls_skills_superior_list[i8]:
                    D9 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1819 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1819 == \
                        ls_skills_superior_list[i9] \
                        or ls_skills_superior_avg_1819 == ls_skills_superior_list[i8]:
                    D10 = book.add_format(property_table_color_shallow)
            #

            # 最小值有两个（考虑的情况没有穷尽）
            if (ls_skills_superior_list[i10] == ls_skills_superior_list[i9] and ls_skills_superior_list[i8] ==
                ls_skills_superior_list[i7]) or \
                    (ls_skills_superior_list[i10] != ls_skills_superior_list[i9] and ls_skills_superior_list[i9] !=
                     ls_skills_superior_list[i8] and ls_skills_superior_list[i8] == ls_skills_superior_list[i7]):
                if ls_skills_superior_avg_01 == ls_skills_superior_list[i10] or ls_skills_superior_avg_01 == \
                        ls_skills_superior_list[i9]:
                    D1 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_23 == ls_skills_superior_list[i10] or ls_skills_superior_avg_23 == \
                        ls_skills_superior_list[i9]:
                    D2 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_45 == ls_skills_superior_list[i10] or ls_skills_superior_avg_45 == \
                        ls_skills_superior_list[i9]:
                    D3 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_67 == ls_skills_superior_list[i10] or ls_skills_superior_avg_67 == \
                        ls_skills_superior_list[i9]:
                    D4 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_89 == ls_skills_superior_list[i10] or ls_skills_superior_avg_89 == \
                        ls_skills_superior_list[i9]:
                    D5 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1011 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1011 == \
                        ls_skills_superior_list[i9]:
                    D6 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1213 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1213 == \
                        ls_skills_superior_list[i9]:
                    D7 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1415 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1415 == \
                        ls_skills_superior_list[i9]:
                    D8 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1617 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1617 == \
                        ls_skills_superior_list[i9]:
                    D9 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1819 == ls_skills_superior_list[i10] or ls_skills_superior_avg_1819 == \
                        ls_skills_superior_list[i9]:
                    D10 = book.add_format(property_table_color_shallow)

            # 最小值只有一个
            if (ls_skills_superior_list[i10] != ls_skills_superior_list[i9] and
                                                    ls_skills_superior_list[i9] == ls_skills_superior_list[i8] and
                                                    ls_skills_superior_list[i9] == ls_skills_superior_list[i7]):
                if ls_skills_superior_avg_01 == ls_skills_superior_list[i10]:
                    D1 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_23 == ls_skills_superior_list[i10]:
                    D2 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_45 == ls_skills_superior_list[i10]:
                    D3 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_67 == ls_skills_superior_list[i10]:
                    D4 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_89 == ls_skills_superior_list[i10]:
                    D5 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1011 == ls_skills_superior_list[i10]:
                    D6 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1213 == ls_skills_superior_list[i10]:
                    D7 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1415 == ls_skills_superior_list[i10]:
                    D8 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1617 == ls_skills_superior_list[i10]:
                    D9 = book.add_format(property_table_color_shallow)
                if ls_skills_superior_avg_1819 == ls_skills_superior_list[i10]:
                    D10 = book.add_format(property_table_color_shallow)

            # 无最小
            if (ls_skills_superior_list[i10] == ls_skills_superior_list[i9] and
                    ls_skills_superior_list[i9] == ls_skills_superior_list[i8] and
                    ls_skills_superior_list[i8] == ls_skills_superior_list[i7]):
                D1 = book.add_format(property_table_color_blank)
                D2 = book.add_format(property_table_color_blank)
                D3 = book.add_format(property_table_color_blank)
                D4 = book.add_format(property_table_color_blank)
                D5 = book.add_format(property_table_color_blank)
                D6 = book.add_format(property_table_color_blank)
                D7 = book.add_format(property_table_color_blank)
                D8 = book.add_format(property_table_color_blank)
                D9 = book.add_format(property_table_color_blank)
                D10 = book.add_format(property_table_color_blank)

            if ls_skills_superior_avg_01 == 0:
                ls_skills_superior_avg_01 = '/'
            sheet.merge_range('F180:F181', ls_skills_superior_avg_01, D1)
            if ls_skills_superior_avg_23 == 0:
                ls_skills_superior_avg_23 = '/'
            sheet.merge_range('F182:F183', ls_skills_superior_avg_23, D2)
            if ls_skills_superior_avg_45 == 0:
                ls_skills_superior_avg_45 = '/'
            sheet.merge_range('F184:F185', ls_skills_superior_avg_45, D3)
            if ls_skills_superior_avg_67 == 0:
                ls_skills_superior_avg_67 = '/'
            sheet.merge_range('F186:F187', ls_skills_superior_avg_67, D4)
            if ls_skills_superior_avg_89 == 0:
                ls_skills_superior_avg_89 = '/'
            sheet.merge_range('F188:F189', ls_skills_superior_avg_89, D5)
            if ls_skills_superior_avg_1011 == 0:
                ls_skills_superior_avg_1011 = '/'
            sheet.merge_range('F190:F191', ls_skills_superior_avg_1011, D6)
            if ls_skills_superior_avg_1213 == 0:
                ls_skills_superior_avg_1213 = '/'
            sheet.merge_range('F192:F193', ls_skills_superior_avg_1213, D7)
            if ls_skills_superior_avg_1415 == 0:
                ls_skills_superior_avg_1415 = '/'
            sheet.merge_range('F194:F195', ls_skills_superior_avg_1415, D8)
            if ls_skills_superior_avg_1617 == 0:
                ls_skills_superior_avg_1617 = '/'
            sheet.merge_range('F196:F197', ls_skills_superior_avg_1617, D9)
            if ls_skills_superior_avg_1819 == 0:
                ls_skills_superior_avg_1819 = '/'
            sheet.merge_range('F198:F199', ls_skills_superior_avg_1819, D10)

            # 填数（同级评分）
            # 异常值判断
            ls_skills_companion_avg_01 = []

            if np.isnan(ls_skills_companion_0[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_01.append(ls_skills_companion_0[i])

            if np.isnan(ls_skills_companion_1[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_01.append(ls_skills_companion_1[i])

            if len(ls_skills_companion_avg_01) == 0:
                ls_skills_companion_avg_01 = 0
            else:
                ls_skills_companion_avg_01 = round(
                    sum(ls_skills_companion_avg_01) / len(ls_skills_companion_avg_01), 2)
            # sheet.merge_range('E180:E181', ls_skills_companion_avg_01, cell_format_table)
            # sheet.write(str('D' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_01, cell_format_content)

            # 异常值判断
            ls_skills_companion_avg_23 = []
            if np.isnan(ls_skills_companion_2[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_23.append(ls_skills_companion_2[i])

            if np.isnan(ls_skills_companion_3[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_23.append(ls_skills_companion_3[i])

            if len(ls_skills_companion_avg_23) == 0:
                ls_skills_companion_avg_23 = 0
            else:
                ls_skills_companion_avg_23 = round(
                    sum(ls_skills_companion_avg_23) / len(ls_skills_companion_avg_23), 2)
            # sheet.merge_range('E182:E183', ls_skills_companion_avg_23, cell_format_table)
            # sheet.write(str('D' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_23, cell_format_content)

            # 异常值判断
            ls_skills_companion_avg_45 = []
            if np.isnan(ls_skills_companion_4[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_45.append(ls_skills_companion_4[i])

            if np.isnan(ls_skills_companion_5[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_45.append(ls_skills_companion_5[i])

            if len(ls_skills_companion_avg_45) == 0:
                ls_skills_companion_avg_45 = 0
            else:
                ls_skills_companion_avg_45 = round(
                    sum(ls_skills_companion_avg_45) / len(ls_skills_companion_avg_45), 2)
            # sheet.merge_range('E184:E185', ls_skills_companion_avg_45, cell_format_table)
            # sheet.write(str('E' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_23, cell_format_content)
            # 异常值判断
            ls_skills_companion_avg_67 = []
            if np.isnan(ls_skills_companion_6[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_67.append(ls_skills_companion_6[i])

            if np.isnan(ls_skills_companion_7[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_67.append(ls_skills_companion_7[i])

            if len(ls_skills_companion_avg_67) == 0:
                ls_skills_companion_avg_67 = 0
            else:
                ls_skills_companion_avg_67 = round(
                    sum(ls_skills_companion_avg_67) / len(ls_skills_companion_avg_67), 2)
            # sheet.merge_range('E186:E187', ls_skills_superior_avg_67, cell_format_table)

            # 异常值判断
            ls_skills_companion_avg_89 = []
            if np.isnan(ls_skills_companion_8[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_89.append(ls_skills_companion_8[i])

            if np.isnan(ls_skills_companion_9[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_89.append(ls_skills_companion_9[i])

            if len(ls_skills_companion_avg_89) == 0:
                ls_skills_companion_avg_89 = 0
            else:
                ls_skills_companion_avg_89 = round(
                    sum(ls_skills_companion_avg_89) / len(ls_skills_companion_avg_89), 2)

            # sheet.merge_range('E188:E189', ls_skills_companion_avg_89, cell_format_table)
            # 异常值判断
            ls_skills_companion_avg_1011 = []
            if np.isnan(ls_skills_companion_10[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1011.append(ls_skills_companion_10[i])

            if np.isnan(ls_skills_companion_11[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1011.append(ls_skills_companion_11[i])

            if len(ls_skills_companion_avg_1011) == 0:
                ls_skills_companion_avg_1011 = 0
            else:
                ls_skills_companion_avg_1011 = round(
                    sum(ls_skills_companion_avg_1011) / len(ls_skills_companion_avg_1011), 2)

            # sheet.merge_range('E190:E191', ls_skills_companion_avg_1011, cell_format_table)
            # 异常值判断
            ls_skills_companion_avg_1213 = []
            if np.isnan(ls_skills_companion_12[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1213.append(ls_skills_companion_12[i])

            if np.isnan(ls_skills_companion_13[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1213.append(ls_skills_companion_13[i])

            if len(ls_skills_companion_avg_1213) == 0:
                ls_skills_companion_avg_1213 = 0
            else:
                ls_skills_companion_avg_1213 = round(
                    sum(ls_skills_companion_avg_1213) / len(ls_skills_companion_avg_1213), 2)
            # sheet.merge_range('E192:E193', ls_skills_companion_avg_1213, cell_format_table)

            # 异常值判断
            ls_skills_companion_avg_1415 = []
            if np.isnan(ls_skills_companion_14[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1415.append(ls_skills_companion_14[i])

            if np.isnan(ls_skills_companion_15[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1415.append(ls_skills_companion_15[i])

            if len(ls_skills_companion_avg_1415) == 0:
                ls_skills_companion_avg_1415 = 0
            else:
                ls_skills_companion_avg_1415 = round(
                    sum(ls_skills_companion_avg_1415) / len(ls_skills_companion_avg_1415), 2)
            # sheet.merge_range('E194:E195', ls_skills_companion_avg_1415, cell_format_table)

            # 异常值判断
            ls_skills_companion_avg_1617 = []
            if np.isnan(ls_skills_companion_16[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1617.append(ls_skills_companion_16[i])

            if np.isnan(ls_skills_companion_17[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1617.append(ls_skills_companion_17[i])

            if len(ls_skills_companion_avg_1617) == 0:
                ls_skills_companion_avg_1617 = 0
            else:
                ls_skills_companion_avg_1617 = round(
                    sum(ls_skills_companion_avg_1617) / len(ls_skills_companion_avg_1617), 2)
            # sheet.merge_range('E196:E197', ls_skills_companion_avg_1617, cell_format_table)

            # 异常值判断
            ls_skills_companion_avg_1819 = []
            if np.isnan(ls_skills_companion_18[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1819.append(ls_skills_companion_18[i])

            if np.isnan(ls_skills_companion_19[i]):
                print('NAN')
            else:
                ls_skills_companion_avg_1819.append(ls_skills_companion_19[i])

            if len(ls_skills_companion_avg_1819) == 0:
                ls_skills_companion_avg_1819 = 0
            else:
                ls_skills_companion_avg_1819 = round(
                    sum(ls_skills_companion_avg_1819) / len(ls_skills_companion_avg_1819), 2)
            # sheet.merge_range('E198:E199', ls_skills_companion_avg_1819, cell_format_table)

            # 排序判断(同级评分)
            ls_skills_companion_list_1 = [ls_skills_companion_avg_01, ls_skills_companion_avg_23,
                                          ls_skills_companion_avg_45, ls_skills_companion_avg_67,
                                          ls_skills_companion_avg_89, ls_skills_companion_avg_1011,
                                          ls_skills_companion_avg_1213, ls_skills_companion_avg_1415,
                                          ls_skills_companion_avg_1617, ls_skills_companion_avg_1819]
            # print(ls_skills_companion_list_1)
            ls_skills_companion_list = []
            for ci in ls_skills_companion_list_1:
                if ci == 0:
                    print('NULL')
                else:
                    ls_skills_companion_list.append(ci)
            print('同级评分：' + str(ls_skills_companion_list))
            #
            # format1 = []
            # 从小到大排序
            dic1_sort_companion = np.argsort(ls_skills_companion_list)  # 用索引引用字典  dict[i]
            dsc_len = len(dic1_sort_companion)
            # 索引排列
            # 取索引的后四位（最大的4位）
            # 防止数组长度变动
            i4 = dic1_sort_companion[dsc_len - 4]  # 第4位
            i3 = dic1_sort_companion[dsc_len - 3]  # 第3位
            i2 = dic1_sort_companion[dsc_len - 2]  # 第2位
            i1 = dic1_sort_companion[dsc_len - 1]  # 第1位
            # 取索引的前四位（最小的4位）
            i7 = dic1_sort_companion[3]  # 第4位
            i8 = dic1_sort_companion[2]  # 第3位
            i9 = dic1_sort_companion[1]  # 第2位
            i10 = dic1_sort_companion[0]  # 第1位
            #
            print(ls_skills_companion_list[i1], ls_skills_companion_list[i2], ls_skills_companion_list[i3]
                  , ls_skills_companion_list[i4], ls_skills_companion_list[i7], ls_skills_companion_list[i8]
                  , ls_skills_companion_list[i9], ls_skills_companion_list[i10])
            # 判断前3位和后3位的逻辑分开写
            # 先预置单元格格式为不添加任何颜色
            E1 = book.add_format(property_table_color_blank)
            E2 = book.add_format(property_table_color_blank)
            E3 = book.add_format(property_table_color_blank)
            E4 = book.add_format(property_table_color_blank)
            E5 = book.add_format(property_table_color_blank)
            E6 = book.add_format(property_table_color_blank)
            E7 = book.add_format(property_table_color_blank)
            E8 = book.add_format(property_table_color_blank)
            E9 = book.add_format(property_table_color_blank)
            E10 = book.add_format(property_table_color_blank)

            # 判断位数(前3位）（最大值有三个）
            if (ls_skills_companion_list[i1] == ls_skills_companion_list[i2] and ls_skills_companion_list[i2] ==
                ls_skills_companion_list[i3] and ls_skills_companion_list[i3] != ls_skills_companion_list[i4]) or \
                    (ls_skills_companion_list[i1] != ls_skills_companion_list[i2] and ls_skills_companion_list[i2] ==
                     ls_skills_companion_list[i3] and ls_skills_companion_list[i3] != ls_skills_companion_list[i4]) or \
                    (ls_skills_companion_list[i1] == ls_skills_companion_list[i2] and ls_skills_companion_list[i2] !=
                     ls_skills_companion_list[i3] and ls_skills_companion_list[i3] != ls_skills_companion_list[i4]) \
                    or (ls_skills_companion_list[i1] != ls_skills_companion_list[i2] and ls_skills_companion_list[i2] !=
                        ls_skills_companion_list[i3] and ls_skills_companion_list[i3] != ls_skills_companion_list[i4]):
                # 有最大的前3项,定位到最大单元格，然后标出颜色
                # 确定最大值，如何关联到颜色（考虑用键值对的形式）例如找到最大值x，那么其对应的y可以被赋值
                if ls_skills_companion_avg_01 == ls_skills_companion_list[i1] or ls_skills_companion_avg_01 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_01 == ls_skills_companion_list[i3]:
                    E1 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_23 == ls_skills_companion_list[i1] or ls_skills_companion_avg_23 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_23 == ls_skills_companion_list[i3]:
                    E2 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_45 == ls_skills_companion_list[i1] or ls_skills_companion_avg_45 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_45 == ls_skills_companion_list[i3]:
                    E3 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_67 == ls_skills_companion_list[i1] or ls_skills_companion_avg_67 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_67 == ls_skills_companion_list[i3]:
                    E4 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_89 == ls_skills_companion_list[i1] or ls_skills_companion_avg_89 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_89 == ls_skills_companion_list[i3]:
                    E5 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1011 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1011 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_1011 == ls_skills_companion_list[i3]:
                    E6 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1213 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1213 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_1213 == ls_skills_companion_list[i3]:
                    E7 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1415 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1415 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_1415 == ls_skills_companion_list[i3]:
                    E8 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1617 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1617 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_1617 == ls_skills_companion_list[i3]:
                    E9 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1819 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1819 == \
                        ls_skills_companion_list[i2] or ls_skills_companion_avg_1819 == ls_skills_companion_list[i3]:
                    E10 = book.add_format(property_table_color_deep)

            # 最大值有两个（考虑的情况没有穷尽）
            if (ls_skills_companion_list[i1] == ls_skills_companion_list[i2] and ls_skills_companion_list[i3] == \
                ls_skills_companion_list[i4]) or (ls_skills_companion_list[i1] != ls_skills_companion_list[i2] and
                                                  ls_skills_companion_list[i2] != ls_skills_companion_list[i3] and
                                                  ls_skills_companion_list[i3] == ls_skills_companion_list[i4]):
                if ls_skills_companion_avg_01 == ls_skills_companion_list[i1] or ls_skills_companion_avg_01 == \
                        ls_skills_companion_list[i2]:
                    E1 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_23 == ls_skills_companion_list[i1] or ls_skills_companion_avg_23 == \
                        ls_skills_companion_list[i2]:
                    E2 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_45 == ls_skills_companion_list[i1] or ls_skills_companion_avg_45 == \
                        ls_skills_companion_list[i2]:
                    E3 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_67 == ls_skills_companion_list[i1] or ls_skills_companion_avg_67 == \
                        ls_skills_companion_list[i2]:
                    E4 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_89 == ls_skills_companion_list[i1] or ls_skills_companion_avg_89 == \
                        ls_skills_companion_list[i2]:
                    E5 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1011 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1011 == \
                        ls_skills_companion_list[i2]:
                    E6 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1213 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1213 == \
                        ls_skills_companion_list[i2]:
                    E7 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1415 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1415 == \
                        ls_skills_companion_list[i2]:
                    E8 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1617 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1617 == \
                        ls_skills_companion_list[i2]:
                    E9 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1819 == ls_skills_companion_list[i1] or ls_skills_companion_avg_1819 == \
                        ls_skills_companion_list[i2]:
                    E10 = book.add_format(property_table_color_deep)

            # 最大值只有一个
            if (ls_skills_companion_list[i1] != ls_skills_companion_list[i2] and
                    ls_skills_companion_list[i2] == ls_skills_companion_list[i3] and
                    ls_skills_companion_list[i2] == ls_skills_companion_list[i4]):
                if ls_skills_companion_avg_01 == ls_skills_companion_list[i1]:
                    E1 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_23 == ls_skills_companion_list[i1]:
                    E2 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_45 == ls_skills_companion_list[i1]:
                    E3 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_67 == ls_skills_companion_list[i1]:
                    E4 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_89 == ls_skills_companion_list[i1]:
                    E5 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1011 == ls_skills_companion_list[i1]:
                    E6 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1213 == ls_skills_companion_list[i1]:
                    E7 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1415 == ls_skills_companion_list[i1]:
                    E8 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1617 == ls_skills_companion_list[i1]:
                    E9 = book.add_format(property_table_color_deep)
                if ls_skills_companion_avg_1819 == ls_skills_companion_list[i1]:
                    E10 = book.add_format(property_table_color_deep)

            # 无最大
            if (ls_skills_companion_list[i1] == ls_skills_companion_list[i2] and
                    ls_skills_companion_list[i2] == ls_skills_companion_list[i3] and
                    ls_skills_companion_list[i2] == ls_skills_companion_list[i4]):
                E1 = book.add_format(property_table_color_blank)
                E2 = book.add_format(property_table_color_blank)
                E3 = book.add_format(property_table_color_blank)
                E4 = book.add_format(property_table_color_blank)
                E5 = book.add_format(property_table_color_blank)
                E6 = book.add_format(property_table_color_blank)
                E7 = book.add_format(property_table_color_blank)
                E8 = book.add_format(property_table_color_blank)
                E9 = book.add_format(property_table_color_blank)
                E10 = book.add_format(property_table_color_blank)

            # 后三位判断（最小值有3位）
            if (ls_skills_companion_list[i10] == ls_skills_companion_list[i9] and ls_skills_companion_list[i9] ==
                ls_skills_companion_list[i8] and ls_skills_companion_list[i8] != ls_skills_companion_list[i7]) or \
                    (ls_skills_companion_list[i10] != ls_skills_companion_list[i9] and ls_skills_companion_list[i9] ==
                     ls_skills_companion_list[i8] and
                     ls_skills_companion_list[i8] != ls_skills_companion_list[i7]) or \
                    (ls_skills_companion_list[i10] == ls_skills_companion_list[i9] and ls_skills_companion_list[i9] !=
                     ls_skills_companion_list[i8] and ls_skills_companion_list[i8] != ls_skills_companion_list[i7]) or \
                    (ls_skills_companion_list[i10] != ls_skills_companion_list[i9] and ls_skills_companion_list[i9] !=
                     ls_skills_companion_list[i8] and ls_skills_companion_list[i8] != ls_skills_companion_list[i7]):
                if ls_skills_companion_avg_01 == ls_skills_companion_list[i10] or ls_skills_companion_avg_01 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_01 == ls_skills_companion_list[i8]:
                    E1 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_23 == ls_skills_companion_list[i10] or ls_skills_companion_avg_23 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_23 == ls_skills_companion_list[i8]:
                    E2 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_45 == ls_skills_companion_list[i10] or ls_skills_companion_avg_45 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_45 == ls_skills_companion_list[i8]:
                    E3 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_67 == ls_skills_companion_list[i10] or ls_skills_companion_avg_67 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_67 == ls_skills_companion_list[i8]:
                    E4 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_89 == ls_skills_companion_list[i10] or ls_skills_companion_avg_89 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_89 == ls_skills_companion_list[i8]:
                    E5 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1011 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1011 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_1011 == ls_skills_companion_list[i8]:
                    E6 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1213 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1213 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_1213 == ls_skills_companion_list[i8]:
                    E7 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1415 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1415 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_1415 == ls_skills_companion_list[i8]:
                    E8 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1617 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1617 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_1617 == ls_skills_companion_list[i8]:
                    E9 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1819 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1819 == \
                        ls_skills_companion_list[i9] \
                        or ls_skills_companion_avg_1819 == ls_skills_companion_list[i8]:
                    E10 = book.add_format(property_table_color_shallow)

            # 最小值有两个（考虑的情况没有穷尽）
            if (ls_skills_companion_list[i10] == ls_skills_companion_list[i9] and ls_skills_companion_list[i8] ==
                ls_skills_companion_list[i7]) or \
                    (ls_skills_companion_list[i10] != ls_skills_companion_list[i9] and ls_skills_companion_list[i9] !=
                     ls_skills_companion_list[i8] and ls_skills_companion_list[i8] == ls_skills_companion_list[i7]):
                if ls_skills_companion_avg_01 == ls_skills_companion_list[i10] or ls_skills_companion_avg_01 == \
                        ls_skills_companion_list[i9]:
                    E1 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_23 == ls_skills_companion_list[i10] or ls_skills_companion_avg_23 == \
                        ls_skills_companion_list[i9]:
                    E2 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_45 == ls_skills_companion_list[i10] or ls_skills_companion_avg_45 == \
                        ls_skills_companion_list[i9]:
                    E3 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_67 == ls_skills_companion_list[i10] or ls_skills_companion_avg_67 == \
                        ls_skills_companion_list[i9]:
                    E4 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_89 == ls_skills_companion_list[i10] or ls_skills_companion_avg_89 == \
                        ls_skills_companion_list[i9]:
                    E5 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1011 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1011 == \
                        ls_skills_companion_list[i9]:
                    E6 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1213 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1213 == \
                        ls_skills_companion_list[i9]:
                    E7 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1415 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1415 == \
                        ls_skills_companion_list[i9]:
                    E8 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1617 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1617 == \
                        ls_skills_companion_list[i9]:
                    E9 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1819 == ls_skills_companion_list[i10] or ls_skills_companion_avg_1819 == \
                        ls_skills_companion_list[i9]:
                    E10 = book.add_format(property_table_color_shallow)

            # 最小值只有一个
            if (ls_skills_companion_list[i10] != ls_skills_companion_list[i9] and
                                                     ls_skills_companion_list[i9] == ls_skills_companion_list[i8] and
                                                     ls_skills_companion_list[i9] == ls_skills_companion_list[i7]):
                if ls_skills_companion_avg_01 == ls_skills_companion_list[i10]:
                    E1 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_23 == ls_skills_companion_list[i10]:
                    E2 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_45 == ls_skills_companion_list[i10]:
                    E3 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_67 == ls_skills_companion_list[i10]:
                    E4 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_89 == ls_skills_companion_list[i10]:
                    E5 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1011 == ls_skills_companion_list[i10]:
                    E6 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1213 == ls_skills_companion_list[i10]:
                    E7 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1415 == ls_skills_companion_list[i10]:
                    E8 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1617 == ls_skills_companion_list[i10]:
                    E9 = book.add_format(property_table_color_shallow)
                if ls_skills_companion_avg_1819 == ls_skills_companion_list[i10]:
                    E10 = book.add_format(property_table_color_shallow)
            # 无最小
            if (ls_skills_companion_list[i10] == ls_skills_companion_list[i9] and
                    ls_skills_companion_list[i9] == ls_skills_companion_list[i8] and
                    ls_skills_companion_list[i8] == ls_skills_companion_list[i7]):
                E1 = book.add_format(property_table_color_blank)
                E2 = book.add_format(property_table_color_blank)
                E3 = book.add_format(property_table_color_blank)
                E4 = book.add_format(property_table_color_blank)
                E5 = book.add_format(property_table_color_blank)
                E6 = book.add_format(property_table_color_blank)
                E7 = book.add_format(property_table_color_blank)
                E8 = book.add_format(property_table_color_blank)
                E9 = book.add_format(property_table_color_blank)
                E10 = book.add_format(property_table_color_blank)

            # 写入数据
            if ls_skills_companion_avg_01 == 0:
                ls_skills_companion_avg_01 = '  /'
            sheet.merge_range('G180:G181', ls_skills_companion_avg_01, E1)
            if ls_skills_companion_avg_23 == 0:
                ls_skills_companion_avg_23 = '  /'
            sheet.merge_range('G182:G183', ls_skills_companion_avg_23, E2)
            if ls_skills_companion_avg_45 == 0:
                ls_skills_companion_avg_45 = '  /'
            sheet.merge_range('G184:G185', ls_skills_companion_avg_45, E3)
            if ls_skills_companion_avg_67 == 0:
                ls_skills_companion_avg_67 = '  /'
            sheet.merge_range('G186:G187', ls_skills_companion_avg_67, E4)
            if ls_skills_companion_avg_89 == 0:
                ls_skills_companion_avg_89 = '  /'
            sheet.merge_range('G188:G189', ls_skills_companion_avg_89, E5)
            if ls_skills_companion_avg_1011 == 0:
                ls_skills_companion_avg_1011 = '  /'
            sheet.merge_range('G190:G191', ls_skills_companion_avg_1011, E6)
            if ls_skills_companion_avg_1213 == 0:
                ls_skills_companion_avg_1213 = '  /'
            sheet.merge_range('G192:G193', ls_skills_companion_avg_1213, E7)
            if ls_skills_companion_avg_1415 == 0:
                ls_skills_companion_avg_1415 = '  /'
            sheet.merge_range('G194:G195', ls_skills_companion_avg_1415, E8)
            if ls_skills_companion_avg_1617 == 0:
                ls_skills_companion_avg_1617 = '  /'
            sheet.merge_range('G196:G197', ls_skills_companion_avg_1617, E9)
            if ls_skills_companion_avg_1819 == 0:
                ls_skills_companion_avg_1819 = '  /'
            sheet.merge_range('G198:G199', ls_skills_companion_avg_1819, E10)

            # 填数（下级评分）
            ls_skills_lower_avg_01 = []
            if np.isnan(ls_skills_lower_0[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_01.append(ls_skills_lower_0[i])

            if np.isnan(ls_skills_lower_1[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_01.append(ls_skills_lower_1[i])

            if len(ls_skills_lower_avg_01) == 0:
                ls_skills_lower_avg_01 = 0
            else:
                ls_skills_lower_avg_01 = round(
                    sum(ls_skills_lower_avg_01) / len(ls_skills_lower_avg_01), 2)

            # sheet.merge_range('F180:F181', ls_skills_lower_avg_01, cell_format_table)
            # sheet.write(str('D' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_01, cell_format_content)
            ls_skills_lower_avg_23 = []
            if np.isnan(ls_skills_lower_2[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_23.append(ls_skills_lower_2[i])

            if np.isnan(ls_skills_lower_3[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_23.append(ls_skills_lower_3[i])

            if len(ls_skills_lower_avg_23) == 0:
                ls_skills_lower_avg_23 = 0
            else:
                ls_skills_lower_avg_23 = round(
                    sum(ls_skills_lower_avg_23) / len(ls_skills_lower_avg_23), 2)
            # sheet.merge_range('F182:F183', ls_skills_lower_avg_23, cell_format_table)
            # sheet.write(str('D' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_23, cell_format_content)
            ls_skills_lower_avg_45 = []
            if np.isnan(ls_skills_lower_4[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_45.append(ls_skills_lower_4[i])

            if np.isnan(ls_skills_lower_5[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_45.append(ls_skills_lower_5[i])

            if len(ls_skills_lower_avg_45) == 0:
                ls_skills_lower_avg_45 = 0
            else:
                ls_skills_lower_avg_45 = round(
                    sum(ls_skills_lower_avg_45) / len(ls_skills_lower_avg_45), 2)
            # sheet.merge_range('F184:F185', ls_skills_lower_avg_45, cell_format_table)
            # sheet.write(str('E' + str(loct3 + len_mo_list2 + 20)), ls_skills_superior_avg_23, cell_format_content)
            ls_skills_lower_avg_67 = []
            if np.isnan(ls_skills_lower_6[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_67.append(ls_skills_lower_6[i])

            if np.isnan(ls_skills_lower_7[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_67.append(ls_skills_lower_7[i])

            if len(ls_skills_lower_avg_67) == 0:
                ls_skills_lower_avg_67 = 0
            else:
                ls_skills_lower_avg_67 = round(
                    sum(ls_skills_lower_avg_67) / len(ls_skills_lower_avg_67), 2)

            # sheet.merge_range('F186:F187', ls_skills_lower_avg_67, cell_format_table)
            ls_skills_lower_avg_89 = []
            if np.isnan(ls_skills_lower_8[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_89.append(ls_skills_lower_8[i])

            if np.isnan(ls_skills_lower_9[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_89.append(ls_skills_lower_9[i])

            if len(ls_skills_lower_avg_89) == 0:
                ls_skills_lower_avg_89 = 0
            else:
                ls_skills_lower_avg_89 = round(
                    sum(ls_skills_lower_avg_89) / len(ls_skills_lower_avg_89), 2)
            # sheet.merge_range('F188:F189', ls_skills_lower_avg_89, cell_format_table)
            ls_skills_lower_avg_1011 = []
            if np.isnan(ls_skills_lower_10[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1011.append(ls_skills_lower_10[i])

            if np.isnan(ls_skills_lower_11[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1011.append(ls_skills_lower_11[i])

            if len(ls_skills_lower_avg_1011) == 0:
                ls_skills_lower_avg_1011 = 0
            else:
                ls_skills_lower_avg_1011 = round(
                    sum(ls_skills_lower_avg_1011) / len(ls_skills_lower_avg_1011), 2)
            # sheet.merge_range('F190:F191', ls_skills_lower_avg_1011, cell_format_table)
            ls_skills_lower_avg_1213 = []
            if np.isnan(ls_skills_lower_12[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1213.append(ls_skills_lower_12[i])

            if np.isnan(ls_skills_lower_13[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1213.append(ls_skills_lower_13[i])

            if len(ls_skills_lower_avg_1213) == 0:
                ls_skills_lower_avg_1213 = 0
            else:
                ls_skills_lower_avg_1213 = round(
                    sum(ls_skills_lower_avg_1213) / len(ls_skills_lower_avg_1213), 2)
            # sheet.merge_range('F192:F193', ls_skills_lower_avg_1213, cell_format_table)
            ls_skills_lower_avg_1415 = []
            if np.isnan(ls_skills_lower_14[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1415.append(ls_skills_lower_14[i])

            if np.isnan(ls_skills_lower_15[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1415.append(ls_skills_lower_15[i])

            if len(ls_skills_lower_avg_1415) == 0:
                ls_skills_lower_avg_1415 = 0
            else:
                ls_skills_lower_avg_1415 = round(
                    sum(ls_skills_lower_avg_1415) / len(ls_skills_lower_avg_1415), 2)
            # sheet.merge_range('F194:F195', ls_skills_lower_avg_1415, cell_format_table)
            ls_skills_lower_avg_1617 = []
            if np.isnan(ls_skills_lower_16[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1617.append(ls_skills_lower_16[i])

            if np.isnan(ls_skills_lower_17[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1617.append(ls_skills_lower_17[i])

            if len(ls_skills_lower_avg_1617) == 0:
                ls_skills_lower_avg_1617 = 0
            else:
                ls_skills_lower_avg_1617 = round(
                    sum(ls_skills_lower_avg_1617) / len(ls_skills_lower_avg_1617), 2)
            # sheet.merge_range('F196:F197', ls_skills_lower_avg_1617, cell_format_table)
            ls_skills_lower_avg_1819 = []
            if np.isnan(ls_skills_lower_18[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1819.append(ls_skills_lower_18[i])

            if np.isnan(ls_skills_lower_19[i]):
                print('NAN')
            else:
                ls_skills_lower_avg_1819.append(ls_skills_lower_19[i])

            if len(ls_skills_lower_avg_1819) == 0:
                ls_skills_lower_avg_1819 = 0
            else:
                ls_skills_lower_avg_1819 = round(
                    sum(ls_skills_lower_avg_1819) / len(ls_skills_lower_avg_1819), 2)
            # sheet.merge_range('F198:F199', ls_skills_lower_avg_1819, cell_format_table)
            #

            # 排序判断(下级评分)
            ls_skills_lower_list_1 = [ls_skills_lower_avg_01, ls_skills_lower_avg_23,
                                      ls_skills_lower_avg_45, ls_skills_lower_avg_67,
                                      ls_skills_lower_avg_89, ls_skills_lower_avg_1011,
                                      ls_skills_lower_avg_1213, ls_skills_lower_avg_1415,
                                      ls_skills_lower_avg_1617, ls_skills_lower_avg_1819]
            ls_skills_lower_list = []
            for li in ls_skills_lower_list_1:
                if li == 0:
                    print('NULL')
                else:
                    ls_skills_lower_list.append(li)
            print('下级评分:' + str(ls_skills_lower_list))

            #
            # format1 = []
            # 从小到大排序
            dic1_sort_lower = np.argsort(ls_skills_lower_list)  # 用索引引用字典  dict[i]
            dsl_len = len(dic1_sort_lower)
            # 索引排列
            # 取索引的后四位（最大的4位）
            i4 = dic1_sort_lower[dsl_len - 4]  # 第4位
            i3 = dic1_sort_lower[dsl_len - 3]  # 第3位
            i2 = dic1_sort_lower[dsl_len - 2]  # 第2位
            i1 = dic1_sort_lower[dsl_len - 1]  # 第1位
            # 取索引的前四位（最小的4位）
            i7 = dic1_sort_lower[3]  # 第4位
            i8 = dic1_sort_lower[2]  # 第3位
            i9 = dic1_sort_lower[1]  # 第2位
            i10 = dic1_sort_lower[0]  # 第1位
            #
            # 判断前3位和后3位的逻辑分开写
            # 先预置单元格格式为不添加任何颜色
            F1 = book.add_format(property_table_color_blank)
            F2 = book.add_format(property_table_color_blank)
            F3 = book.add_format(property_table_color_blank)
            F4 = book.add_format(property_table_color_blank)
            F5 = book.add_format(property_table_color_blank)
            F6 = book.add_format(property_table_color_blank)
            F7 = book.add_format(property_table_color_blank)
            F8 = book.add_format(property_table_color_blank)
            F9 = book.add_format(property_table_color_blank)
            F10 = book.add_format(property_table_color_blank)

            # 判断位数(前3位）（最大值有三个）
            if (ls_skills_lower_list[i1] == ls_skills_lower_list[i2] and ls_skills_lower_list[i2] ==
                ls_skills_lower_list[i3] and ls_skills_lower_list[i3] != ls_skills_lower_list[i4]) or \
                    (ls_skills_lower_list[i1] != ls_skills_lower_list[i2] and ls_skills_lower_list[i2] ==
                     ls_skills_lower_list[i3] and ls_skills_lower_list[i3] != ls_skills_lower_list[i4]) or \
                    (ls_skills_lower_list[i1] == ls_skills_lower_list[i2] and ls_skills_lower_list[i2] !=
                     ls_skills_lower_list[i3] and ls_skills_lower_list[i3] != ls_skills_lower_list[i4]) \
                    or (ls_skills_lower_list[i1] != ls_skills_lower_list[i2] and ls_skills_lower_list[i2] !=
                        ls_skills_lower_list[i3] and ls_skills_lower_list[i3] != ls_skills_lower_list[i4]):
                # 有最大的前3项,定位到最大单元格，然后标出颜色
                # 确定最大值，如何关联到颜色（考虑用键值对的形式）例如找到最大值x，那么其对应的y可以被赋值
                if ls_skills_lower_avg_01 == ls_skills_lower_list[i1] or ls_skills_lower_avg_01 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_01 == ls_skills_lower_list[i3]:
                    F1 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_23 == ls_skills_lower_list[i1] or ls_skills_lower_avg_23 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_23 == ls_skills_lower_list[i3]:
                    F2 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_45 == ls_skills_lower_list[i1] or ls_skills_lower_avg_45 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_45 == ls_skills_lower_list[i3]:
                    F3 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_67 == ls_skills_lower_list[i1] or ls_skills_lower_avg_67 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_67 == ls_skills_lower_list[i3]:
                    F4 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_89 == ls_skills_lower_list[i1] or ls_skills_lower_avg_89 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_89 == ls_skills_lower_list[i3]:
                    F5 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1011 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1011 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_1011 == ls_skills_lower_list[i3]:
                    F6 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1213 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1213 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_1213 == ls_skills_lower_list[i3]:
                    F7 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1415 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1415 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_1415 == ls_skills_lower_list[i3]:
                    F8 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1617 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1617 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_1617 == ls_skills_lower_list[i3]:
                    F9 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1819 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1819 == \
                        ls_skills_lower_list[i2] or ls_skills_lower_avg_1819 == ls_skills_lower_list[i3]:
                    F10 = book.add_format(property_table_color_deep)

            # 最大值有两个（考虑的情况没有穷尽）
            if (ls_skills_lower_list[i1] == ls_skills_lower_list[i2] and ls_skills_lower_list[i3] == \
                ls_skills_lower_list[i4]) or (ls_skills_lower_list[i1] != ls_skills_lower_list[i2] and
                                              ls_skills_lower_list[i2] != ls_skills_lower_list[i3] and
                                              ls_skills_lower_list[i3] == ls_skills_lower_list[i4]):
                if ls_skills_lower_avg_01 == ls_skills_lower_list[i1] or ls_skills_lower_avg_01 == \
                        ls_skills_lower_list[i2]:
                    F1 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_23 == ls_skills_lower_list[i1] or ls_skills_lower_avg_23 == \
                        ls_skills_lower_list[i2]:
                    F2 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_45 == ls_skills_lower_list[i1] or ls_skills_lower_avg_45 == \
                        ls_skills_lower_list[i2]:
                    F3 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_67 == ls_skills_lower_list[i1] or ls_skills_lower_avg_67 == \
                        ls_skills_lower_list[i2]:
                    F4 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_89 == ls_skills_lower_list[i1] or ls_skills_lower_avg_89 == \
                        ls_skills_lower_list[i2]:
                    F5 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1011 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1011 == \
                        ls_skills_lower_list[i2]:
                    F6 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1213 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1213 == \
                        ls_skills_lower_list[i2]:
                    F7 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1415 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1415 == \
                        ls_skills_lower_list[i2]:
                    F8 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1617 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1617 == \
                        ls_skills_lower_list[i2]:
                    F9 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1819 == ls_skills_lower_list[i1] or ls_skills_lower_avg_1819 == \
                        ls_skills_lower_list[i2]:
                    F10 = book.add_format(property_table_color_deep)

            # 最大值只有一个
            if (ls_skills_lower_list[i1] != ls_skills_lower_list[i2] and
                    ls_skills_lower_list[i2] == ls_skills_lower_list[i3] and
                    ls_skills_lower_list[i2] == ls_skills_lower_list[i4]):
                if ls_skills_lower_avg_01 == ls_skills_lower_list[i1]:
                    F1 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_23 == ls_skills_lower_list[i1]:
                    F2 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_45 == ls_skills_lower_list[i1]:
                    F3 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_67 == ls_skills_lower_list[i1]:
                    F4 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_89 == ls_skills_lower_list[i1]:
                    F5 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1011 == ls_skills_lower_list[i1]:
                    F6 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1213 == ls_skills_lower_list[i1]:
                    F7 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1415 == ls_skills_lower_list[i1]:
                    F8 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1617 == ls_skills_lower_list[i1]:
                    F9 = book.add_format(property_table_color_deep)
                if ls_skills_lower_avg_1819 == ls_skills_lower_list[i1]:
                    F10 = book.add_format(property_table_color_deep)

            # 无最大
            if (ls_skills_lower_list[i1] == ls_skills_lower_list[i2] and
                    ls_skills_lower_list[i2] == ls_skills_lower_list[i3] and
                    ls_skills_lower_list[i3] == ls_skills_lower_list[i4]):
                F1 = book.add_format(property_table_color_blank)
                F2 = book.add_format(property_table_color_blank)
                F3 = book.add_format(property_table_color_blank)
                F4 = book.add_format(property_table_color_blank)
                F5 = book.add_format(property_table_color_blank)
                F6 = book.add_format(property_table_color_blank)
                F7 = book.add_format(property_table_color_blank)
                F8 = book.add_format(property_table_color_blank)
                F9 = book.add_format(property_table_color_blank)
                F10 = book.add_format(property_table_color_blank)

            # 后三位判断（最小值有3位）
            if (ls_skills_lower_list[i10] == ls_skills_lower_list[i9] and ls_skills_lower_list[i9] ==
                ls_skills_lower_list[i8] and ls_skills_lower_list[i8] != ls_skills_lower_list[i7]) or \
                    (ls_skills_lower_list[i10] != ls_skills_lower_list[i9] and ls_skills_lower_list[i9] ==
                     ls_skills_lower_list[i8] and
                     ls_skills_lower_list[i8] != ls_skills_lower_list[i7]) or \
                    (ls_skills_lower_list[i10] == ls_skills_lower_list[i9] and ls_skills_lower_list[i9] !=
                     ls_skills_lower_list[i8] and ls_skills_lower_list[i8] != ls_skills_lower_list[i7]) or \
                    (ls_skills_lower_list[i10] != ls_skills_lower_list[i9] and ls_skills_lower_list[i9] !=
                     ls_skills_lower_list[i8] and ls_skills_lower_list[i8] != ls_skills_lower_list[i7]):
                if ls_skills_lower_avg_01 == ls_skills_lower_list[i10] or ls_skills_lower_avg_01 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_01 == ls_skills_lower_list[i8]:
                    F1 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_23 == ls_skills_lower_list[i10] or ls_skills_lower_avg_23 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_23 == ls_skills_lower_list[i8]:
                    F2 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_45 == ls_skills_lower_list[i10] or ls_skills_lower_avg_45 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_45 == ls_skills_lower_list[i8]:
                    F3 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_67 == ls_skills_lower_list[i10] or ls_skills_lower_avg_67 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_67 == ls_skills_lower_list[i8]:
                    F4 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_89 == ls_skills_lower_list[i10] or ls_skills_lower_avg_89 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_89 == ls_skills_lower_list[i8]:
                    F5 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1011 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1011 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_1011 == ls_skills_lower_list[i8]:
                    F6 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1213 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1213 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_1213 == ls_skills_lower_list[i8]:
                    F7 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1415 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1415 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_1415 == ls_skills_lower_list[i8]:
                    F8 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1617 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1617 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_1617 == ls_skills_lower_list[i8]:
                    F9 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1819 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1819 == \
                        ls_skills_lower_list[i9] \
                        or ls_skills_lower_avg_1819 == ls_skills_lower_list[i8]:
                    F10 = book.add_format(property_table_color_shallow)

            # 最小值有两个（考虑的情况没有穷尽）
            if (ls_skills_lower_list[i10] == ls_skills_lower_list[i9] and ls_skills_lower_list[i8] ==
                ls_skills_lower_list[i7]) or \
                    (ls_skills_lower_list[i10] != ls_skills_lower_list[i9] and ls_skills_lower_list[i9] !=
                     ls_skills_lower_list[i8] and ls_skills_lower_list[i8] == ls_skills_lower_list[i7]):
                if ls_skills_lower_avg_01 == ls_skills_lower_list[i10] or ls_skills_lower_avg_01 == \
                        ls_skills_lower_list[i9]:
                    F1 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_23 == ls_skills_lower_list[i10] or ls_skills_lower_avg_23 == \
                        ls_skills_lower_list[i9]:
                    F2 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_45 == ls_skills_lower_list[i10] or ls_skills_lower_avg_45 == \
                        ls_skills_lower_list[i9]:
                    F3 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_67 == ls_skills_lower_list[i10] or ls_skills_lower_avg_67 == \
                        ls_skills_lower_list[i9]:
                    F4 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_89 == ls_skills_lower_list[i10] or ls_skills_lower_avg_89 == \
                        ls_skills_lower_list[i9]:
                    F5 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1011 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1011 == \
                        ls_skills_lower_list[i9]:
                    F6 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1213 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1213 == \
                        ls_skills_lower_list[i9]:
                    F7 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1415 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1415 == \
                        ls_skills_lower_list[i9]:
                    F8 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1617 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1617 == \
                        ls_skills_lower_list[i9]:
                    F9 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1819 == ls_skills_lower_list[i10] or ls_skills_lower_avg_1819 == \
                        ls_skills_lower_list[i9]:
                    F10 = book.add_format(property_table_color_shallow)

            # 最小值只有一个
            if (ls_skills_lower_list[i10] != ls_skills_lower_list[i9] and
                    ls_skills_lower_list[i9] == ls_skills_lower_list[i8] and
                    ls_skills_lower_list[i9] == ls_skills_lower_list[i7]):
                if ls_skills_lower_avg_01 == ls_skills_lower_list[i10]:
                    F1 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_23 == ls_skills_lower_list[i10]:
                    F2 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_45 == ls_skills_lower_list[i10]:
                    F3 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_67 == ls_skills_lower_list[i10]:
                    F4 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_89 == ls_skills_lower_list[i10]:
                    F5 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1011 == ls_skills_lower_list[i10]:
                    F6 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1213 == ls_skills_lower_list[i10]:
                    F7 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1415 == ls_skills_lower_list[i10]:
                    F8 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1617 == ls_skills_lower_list[i10]:
                    F9 = book.add_format(property_table_color_shallow)
                if ls_skills_lower_avg_1819 == ls_skills_lower_list[i10]:
                    F10 = book.add_format(property_table_color_shallow)

            # 无最大
            if (ls_skills_lower_list[i10] == ls_skills_lower_list[i9] and
                    ls_skills_lower_list[i9] == ls_skills_lower_list[i8] and
                    ls_skills_lower_list[i9] == ls_skills_lower_list[i7]):
                F1 = book.add_format(property_table_color_blank)
                F2 = book.add_format(property_table_color_blank)
                F3 = book.add_format(property_table_color_blank)
                F4 = book.add_format(property_table_color_blank)
                F5 = book.add_format(property_table_color_blank)
                F6 = book.add_format(property_table_color_blank)
                F7 = book.add_format(property_table_color_blank)
                F8 = book.add_format(property_table_color_blank)
                F9 = book.add_format(property_table_color_blank)
                F10 = book.add_format(property_table_color_blank)
            #
            if ls_skills_lower_avg_01 == 0:
                ls_skills_lower_avg_01 = '/'
            sheet.merge_range('H180:H181', ls_skills_lower_avg_01, F1)
            if ls_skills_lower_avg_23 == 0:
                ls_skills_lower_avg_23 = '/'
            sheet.merge_range('H182:H183', ls_skills_lower_avg_23, F2)
            if ls_skills_lower_avg_45 == 0:
                ls_skills_lower_avg_45 = '/'
            sheet.merge_range('H184:H185', ls_skills_lower_avg_45, F3)
            if ls_skills_lower_avg_67 == 0:
                ls_skills_lower_avg_67 = '/'
            sheet.merge_range('H186:H187', ls_skills_lower_avg_67, F4)
            if ls_skills_lower_avg_89 == 0:
                ls_skills_lower_avg_89 = '/'
            sheet.merge_range('H188:H189', ls_skills_lower_avg_89, F5)
            if ls_skills_lower_avg_1011 == 0:
                ls_skills_lower_avg_1011 = '/'
            sheet.merge_range('H190:H191', ls_skills_lower_avg_1011, F6)
            if ls_skills_lower_avg_1213 == 0:
                ls_skills_lower_avg_1213 = '/'
            sheet.merge_range('H192:H193', ls_skills_lower_avg_1213, F7)
            if ls_skills_lower_avg_1415 == 0:
                ls_skills_lower_avg_1415 = '/'
            sheet.merge_range('H194:H195', ls_skills_lower_avg_1415, F8)
            if ls_skills_lower_avg_1617 == 0:
                ls_skills_lower_avg_1617 = '/'
            sheet.merge_range('H196:H197', ls_skills_lower_avg_1617, F9)
            if ls_skills_lower_avg_1819 == 0:
                ls_skills_lower_avg_1819 = '/'
            sheet.merge_range('H198:H199', ls_skills_lower_avg_1819, F10)

            # 填数（他评均分）
            # 01
            ls_skills_other_avg_01 = []
            if np.isnan(ls_skills_other_0[i]):
                print('NAN')
            else:
                ls_skills_other_avg_01.append(ls_skills_other_0[i])

            if np.isnan(ls_skills_other_1[i]):
                print('NAN')
            else:
                ls_skills_other_avg_01.append(ls_skills_other_1[i])

            if len(ls_skills_other_avg_01) == 0:
                ls_skills_other_avg_01 = 0
            else:
                ls_skills_other_avg_01 = round(
                    sum(ls_skills_other_avg_01) / len(ls_skills_other_avg_01), 2)
            # 23
            ls_skills_other_avg_23 = []
            if np.isnan(ls_skills_other_2[i]):
                print('NAN')
            else:
                ls_skills_other_avg_23.append(ls_skills_other_2[i])

            if np.isnan(ls_skills_other_3[i]):
                print('NAN')
            else:
                ls_skills_other_avg_23.append(ls_skills_other_3[i])

            if len(ls_skills_other_avg_23) == 0:
                ls_skills_other_avg_23 = 0
            else:
                ls_skills_other_avg_23 = round(
                    sum(ls_skills_other_avg_23) / len(ls_skills_other_avg_23), 2)
            # 45
            ls_skills_other_avg_45 = []
            if np.isnan(ls_skills_other_4[i]):
                print('NAN')
            else:
                ls_skills_other_avg_45.append(ls_skills_other_4[i])

            if np.isnan(ls_skills_other_5[i]):
                print('NAN')
            else:
                ls_skills_other_avg_45.append(ls_skills_other_5[i])

            if len(ls_skills_other_avg_45) == 0:
                ls_skills_other_avg_45 = 0
            else:
                ls_skills_other_avg_45 = round(
                    sum(ls_skills_other_avg_45) / len(ls_skills_other_avg_45), 2)
            # 67
            ls_skills_other_avg_67 = []
            if np.isnan(ls_skills_other_6[i]):
                print('NAN')
            else:
                ls_skills_other_avg_67.append(ls_skills_other_6[i])

            if np.isnan(ls_skills_other_7[i]):
                print('NAN')
            else:
                ls_skills_other_avg_67.append(ls_skills_other_7[i])

            if len(ls_skills_other_avg_67) == 0:
                ls_skills_other_avg_67 = 0
            else:
                ls_skills_other_avg_67 = round(
                    sum(ls_skills_other_avg_67) / len(ls_skills_other_avg_67), 2)
            # 89
            ls_skills_other_avg_89 = []
            if np.isnan(ls_skills_other_8[i]):
                print('NAN')
            else:
                ls_skills_other_avg_89.append(ls_skills_other_8[i])

            if np.isnan(ls_skills_other_9[i]):
                print('NAN')
            else:
                ls_skills_other_avg_89.append(ls_skills_other_9[i])

            if len(ls_skills_other_avg_89) == 0:
                ls_skills_other_avg_89 = 0
            else:
                ls_skills_other_avg_89 = round(
                    sum(ls_skills_other_avg_89) / len(ls_skills_other_avg_89), 2)
            # 1011
            ls_skills_other_avg_1011 = []
            if np.isnan(ls_skills_other_10[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1011.append(ls_skills_other_10[i])

            if np.isnan(ls_skills_other_11[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1011.append(ls_skills_other_11[i])

            if len(ls_skills_other_avg_1011) == 0:
                ls_skills_other_avg_1011 = 0
            else:
                ls_skills_other_avg_1011 = round(
                    sum(ls_skills_other_avg_1011) / len(ls_skills_other_avg_1011), 2)
            # 1213
            ls_skills_other_avg_1213 = []
            if np.isnan(ls_skills_other_12[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1213.append(ls_skills_other_12[i])

            if np.isnan(ls_skills_other_13[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1213.append(ls_skills_other_13[i])

            if len(ls_skills_other_avg_1213) == 0:
                ls_skills_other_avg_1213 = 0
            else:
                ls_skills_other_avg_1213 = round(
                    sum(ls_skills_other_avg_1213) / len(ls_skills_other_avg_1213), 2)
            # 1415
            ls_skills_other_avg_1415 = []
            if np.isnan(ls_skills_other_14[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1415.append(ls_skills_other_14[i])

            if np.isnan(ls_skills_other_15[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1415.append(ls_skills_other_15[i])

            if len(ls_skills_other_avg_1415) == 0:
                ls_skills_other_avg_1415 = 0
            else:
                ls_skills_other_avg_1415 = round(
                    sum(ls_skills_other_avg_1415) / len(ls_skills_other_avg_1415), 2)
            # 1617
            ls_skills_other_avg_1617 = []
            if np.isnan(ls_skills_other_16[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1617.append(ls_skills_other_16[i])

            if np.isnan(ls_skills_other_17[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1617.append(ls_skills_other_17[i])

            if len(ls_skills_other_avg_1617) == 0:
                ls_skills_other_avg_1617 = 0
            else:
                ls_skills_other_avg_1617 = round(
                    sum(ls_skills_other_avg_1617) / len(ls_skills_other_avg_1617), 2)

            # 1819
            ls_skills_other_avg_1819 = []
            if np.isnan(ls_skills_other_18[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1819.append(ls_skills_other_18[i])

            if np.isnan(ls_skills_other_19[i]):
                print('NAN')
            else:
                ls_skills_other_avg_1819.append(ls_skills_other_19[i])

            if len(ls_skills_other_avg_1819) == 0:
                ls_skills_other_avg_1819 = 0
            else:
                ls_skills_other_avg_1819 = round(
                    sum(ls_skills_other_avg_1819) / len(ls_skills_other_avg_1819), 2)

            # 排序判断(下级评分)
            ls_skills_other_list_1 = [ls_skills_other_avg_01, ls_skills_other_avg_23,
                                      ls_skills_other_avg_45, ls_skills_other_avg_67,
                                      ls_skills_other_avg_89, ls_skills_other_avg_1011,
                                      ls_skills_other_avg_1213, ls_skills_other_avg_1415,
                                      ls_skills_other_avg_1617, ls_skills_other_avg_1819]
            ls_skills_other_list = []
            for oi in ls_skills_other_list_1:
                if oi == 0:
                    print('NULL')
                else:
                    ls_skills_other_list.append(oi)
            print('他评均分：' + str(ls_skills_other_list))

            #
            # format1 = []
            # 从小到大排序
            dic1_sort_other = np.argsort(ls_skills_other_list)  # 用索引引用字典  dict[i]
            dsc_len = len(dic1_sort_other)
            # 索引排列
            # 取索引的后四位（最大的4位）
            i4 = dic1_sort_other[dsc_len - 4]  # 第4位
            i3 = dic1_sort_other[dsc_len - 3]  # 第3位
            i2 = dic1_sort_other[dsc_len - 2]  # 第2位
            i1 = dic1_sort_other[dsc_len - 1]  # 第1位
            # 取索引的前四位（最小的4位）
            i7 = dic1_sort_other[3]  # 第4位
            i8 = dic1_sort_other[2]  # 第3位
            i9 = dic1_sort_other[1]  # 第2位
            i10 = dic1_sort_other[0]  # 第1位
            #
            # 判断前3位和后3位的逻辑分开写
            # 先预置单元格格式为不添加任何颜色
            G1 = book.add_format(property_table_color_blank)
            G2 = book.add_format(property_table_color_blank)
            G3 = book.add_format(property_table_color_blank)
            G4 = book.add_format(property_table_color_blank)
            G5 = book.add_format(property_table_color_blank)
            G6 = book.add_format(property_table_color_blank)
            G7 = book.add_format(property_table_color_blank)
            G8 = book.add_format(property_table_color_blank)
            G9 = book.add_format(property_table_color_blank)
            G10 = book.add_format(property_table_color_blank)

            # 判断位数(前3位）（最大值有三个）
            if (ls_skills_other_list[i1] == ls_skills_other_list[i2] and ls_skills_other_list[i2] ==
                ls_skills_other_list[i3] and ls_skills_other_list[i3] != ls_skills_other_list[i4]) or \
                    (ls_skills_other_list[i1] != ls_skills_other_list[i2] and ls_skills_other_list[i2] ==
                     ls_skills_other_list[i3] and ls_skills_other_list[i3] != ls_skills_other_list[i4]) or \
                    (ls_skills_other_list[i1] == ls_skills_other_list[i2] and ls_skills_other_list[i2] !=
                     ls_skills_other_list[i3] and ls_skills_other_list[i3] != ls_skills_other_list[i4]) \
                    or (ls_skills_other_list[i1] != ls_skills_other_list[i2] and ls_skills_other_list[i2] !=
                        ls_skills_other_list[i3] and ls_skills_other_list[i3] != ls_skills_other_list[i4]):
                # 有最大的前3项,定位到最大单元格，然后标出颜色
                # 确定最大值，如何关联到颜色（考虑用键值对的形式）例如找到最大值x，那么其对应的y可以被赋值
                if ls_skills_other_avg_01 == ls_skills_other_list[i1] or ls_skills_other_avg_01 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_01 == ls_skills_other_list[i3]:
                    G1 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_23 == ls_skills_other_list[i1] or ls_skills_other_avg_23 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_23 == ls_skills_other_list[i3]:
                    G2 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_45 == ls_skills_other_list[i1] or ls_skills_other_avg_45 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_45 == ls_skills_other_list[i3]:
                    G3 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_67 == ls_skills_other_list[i1] or ls_skills_other_avg_67 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_67 == ls_skills_other_list[i3]:
                    G4 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_89 == ls_skills_other_list[i1] or ls_skills_other_avg_89 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_89 == ls_skills_other_list[i3]:
                    G5 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1011 == ls_skills_other_list[i1] or ls_skills_other_avg_1011 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_1011 == ls_skills_other_list[i3]:
                    G6 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1213 == ls_skills_other_list[i1] or ls_skills_other_avg_1213 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_1213 == ls_skills_other_list[i3]:
                    G7 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1415 == ls_skills_other_list[i1] or ls_skills_other_avg_1415 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_1415 == ls_skills_other_list[i3]:
                    G8 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1617 == ls_skills_other_list[i1] or ls_skills_other_avg_1617 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_1617 == ls_skills_other_list[i3]:
                    G9 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1819 == ls_skills_other_list[i1] or ls_skills_other_avg_1819 == \
                        ls_skills_other_list[i2] or ls_skills_other_avg_1819 == ls_skills_other_list[i3]:
                    G10 = book.add_format(property_table_color_deep)

            # 最大值有两个（考虑的情况没有穷尽）
            if (ls_skills_other_list[i1] == ls_skills_other_list[i2] and ls_skills_other_list[i3] == \
                ls_skills_other_list[i4] and ls_skills_other_list[i2] != ls_skills_other_list[i3]) or \
                    (ls_skills_other_list[i1] != ls_skills_other_list[i2] and
                     ls_skills_other_list[i2] != ls_skills_other_list[i3] and
                     ls_skills_other_list[i3] == ls_skills_other_list[i4]):
                if ls_skills_other_avg_01 == ls_skills_other_list[i1] or ls_skills_other_avg_01 == \
                        ls_skills_other_list[i2]:
                    G1 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_23 == ls_skills_other_list[i1] or ls_skills_other_avg_23 == \
                        ls_skills_other_list[i2]:
                    G2 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_45 == ls_skills_other_list[i1] or ls_skills_other_avg_45 == \
                        ls_skills_other_list[i2]:
                    G3 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_67 == ls_skills_other_list[i1] or ls_skills_other_avg_67 == \
                        ls_skills_other_list[i2]:
                    G4 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_89 == ls_skills_other_list[i1] or ls_skills_other_avg_89 == \
                        ls_skills_other_list[i2]:
                    G5 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1011 == ls_skills_other_list[i1] or ls_skills_other_avg_1011 == \
                        ls_skills_other_list[i2]:
                    G6 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1213 == ls_skills_other_list[i1] or ls_skills_other_avg_1213 == \
                        ls_skills_other_list[i2]:
                    G7 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1415 == ls_skills_other_list[i1] or ls_skills_other_avg_1415 == \
                        ls_skills_other_list[i2]:
                    G8 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1617 == ls_skills_other_list[i1] or ls_skills_other_avg_1617 == \
                        ls_skills_other_list[i2]:
                    G9 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1819 == ls_skills_other_list[i1] or ls_skills_other_avg_1819 == \
                        ls_skills_other_list[i2]:
                    G10 = book.add_format(property_table_color_deep)

            # 最大值只有一个
            if (ls_skills_other_list[i1] != ls_skills_other_list[i2] and
                    ls_skills_other_list[i2] == ls_skills_other_list[i3] and
                    ls_skills_other_list[i2] == ls_skills_other_list[i4]):
                if ls_skills_other_avg_01 == ls_skills_other_list[i1]:
                    G1 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_23 == ls_skills_other_list[i1]:
                    G2 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_45 == ls_skills_other_list[i1]:
                    G3 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_67 == ls_skills_other_list[i1]:
                    G4 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_89 == ls_skills_other_list[i1]:
                    G5 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1011 == ls_skills_other_list[i1]:
                    G6 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1213 == ls_skills_other_list[i1]:
                    G7 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1415 == ls_skills_other_list[i1]:
                    G8 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1617 == ls_skills_other_list[i1]:
                    G9 = book.add_format(property_table_color_deep)
                if ls_skills_other_avg_1819 == ls_skills_other_list[i1]:
                    G10 = book.add_format(property_table_color_deep)

            # 无最大
            if (ls_skills_lower_list[i1] == ls_skills_lower_list[i2] and
                    ls_skills_lower_list[i2] == ls_skills_lower_list[i3] and
                    ls_skills_lower_list[i3] == ls_skills_lower_list[i4]):
                G1 = book.add_format(property_table_color_blank)
                G2 = book.add_format(property_table_color_blank)
                G3 = book.add_format(property_table_color_blank)
                G4 = book.add_format(property_table_color_blank)
                G5 = book.add_format(property_table_color_blank)
                G6 = book.add_format(property_table_color_blank)
                G7 = book.add_format(property_table_color_blank)
                G8 = book.add_format(property_table_color_blank)
                G9 = book.add_format(property_table_color_blank)
                G10 = book.add_format(property_table_color_blank)

            # 后三位判断（最小值有3位）
            if (ls_skills_other_list[i10] == ls_skills_other_list[i9] and ls_skills_other_list[i9] ==
                ls_skills_other_list[i8] and ls_skills_other_list[i8] != ls_skills_other_list[i7]) or \
                    (ls_skills_other_list[i10] != ls_skills_other_list[i9] and ls_skills_other_list[i9] ==
                     ls_skills_other_list[i8] and
                     ls_skills_other_list[i8] != ls_skills_other_list[i7]) or \
                    (ls_skills_other_list[i10] == ls_skills_other_list[i9] and ls_skills_other_list[i9] !=
                     ls_skills_other_list[i8] and ls_skills_other_list[i8] != ls_skills_other_list[i7]) or \
                    (ls_skills_other_list[i10] != ls_skills_other_list[i9] and ls_skills_other_list[i9] !=
                     ls_skills_other_list[i8] and ls_skills_other_list[i8] != ls_skills_other_list[i7]):
                if ls_skills_other_avg_01 == ls_skills_other_list[i10] or ls_skills_other_avg_01 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_01 == ls_skills_other_list[i8]:
                    G1 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_23 == ls_skills_other_list[i10] or ls_skills_other_avg_23 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_23 == ls_skills_other_list[i8]:
                    G2 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_45 == ls_skills_other_list[i10] or ls_skills_other_avg_45 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_45 == ls_skills_other_list[i8]:
                    G3 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_67 == ls_skills_other_list[i10] or ls_skills_other_avg_67 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_67 == ls_skills_other_list[i8]:
                    G4 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_89 == ls_skills_other_list[i10] or ls_skills_other_avg_89 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_89 == ls_skills_other_list[i8]:
                    G5 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1011 == ls_skills_other_list[i10] or ls_skills_other_avg_1011 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_1011 == ls_skills_other_list[i8]:
                    G6 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1213 == ls_skills_other_list[i10] or ls_skills_other_avg_1213 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_1213 == ls_skills_other_list[i8]:
                    G7 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1415 == ls_skills_other_list[i10] or ls_skills_other_avg_1415 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_1415 == ls_skills_other_list[i8]:
                    G8 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1617 == ls_skills_other_list[i10] or ls_skills_other_avg_1617 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_1617 == ls_skills_other_list[i8]:
                    G9 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1819 == ls_skills_other_list[i10] or ls_skills_other_avg_1819 == \
                        ls_skills_other_list[i9] \
                        or ls_skills_other_avg_1819 == ls_skills_other_list[i8]:
                    G10 = book.add_format(property_table_color_shallow)

            # 最小值有两个（考虑的情况没有穷尽）
            if (ls_skills_other_list[i10] == ls_skills_other_list[i9] and ls_skills_other_list[i8] ==
                ls_skills_other_list[i7]) or \
                    (ls_skills_other_list[i10] != ls_skills_other_list[i9] and ls_skills_other_list[i9] !=
                     ls_skills_other_list[i8] and ls_skills_other_list[i8] == ls_skills_other_list[i7]):
                if ls_skills_other_avg_01 == ls_skills_other_list[i10] or ls_skills_other_avg_01 == \
                        ls_skills_other_list[i9]:
                    G1 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_23 == ls_skills_other_list[i10] or ls_skills_other_avg_23 == \
                        ls_skills_other_list[i9]:
                    G2 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_45 == ls_skills_other_list[i10] or ls_skills_other_avg_45 == \
                        ls_skills_other_list[i9]:
                    G3 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_67 == ls_skills_other_list[i10] or ls_skills_other_avg_67 == \
                        ls_skills_other_list[i9]:
                    G4 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_89 == ls_skills_other_list[i10] or ls_skills_other_avg_89 == \
                        ls_skills_other_list[i9]:
                    G5 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1011 == ls_skills_other_list[i10] or ls_skills_other_avg_1011 == \
                        ls_skills_other_list[i9]:
                    G6 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1213 == ls_skills_other_list[i10] or ls_skills_other_avg_1213 == \
                        ls_skills_other_list[i9]:
                    G7 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1415 == ls_skills_other_list[i10] or ls_skills_other_avg_1415 == \
                        ls_skills_other_list[i9]:
                    G8 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1617 == ls_skills_other_list[i10] or ls_skills_other_avg_1617 == \
                        ls_skills_other_list[i9]:
                    G9 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1819 == ls_skills_other_list[i10] or ls_skills_other_avg_1819 == \
                        ls_skills_other_list[i9]:
                    G10 = book.add_format(property_table_color_shallow)

            # 最小值只有一个
            if (ls_skills_other_list[i10] != ls_skills_other_list[i9] and
                    ls_skills_other_list[i9] == ls_skills_other_list[i8] and
                    ls_skills_other_list[i9] == ls_skills_other_list[i7]):
                if ls_skills_other_avg_01 == ls_skills_other_list[i10]:
                    G1 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_23 == ls_skills_other_list[i10]:
                    G2 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_45 == ls_skills_other_list[i10]:
                    G3 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_67 == ls_skills_other_list[i10]:
                    G4 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_89 == ls_skills_other_list[i10]:
                    G5 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1011 == ls_skills_other_list[i10]:
                    G6 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1213 == ls_skills_other_list[i10]:
                    G7 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1415 == ls_skills_other_list[i10]:
                    G8 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1617 == ls_skills_other_list[i10]:
                    G9 = book.add_format(property_table_color_shallow)
                if ls_skills_other_avg_1819 == ls_skills_other_list[i10]:
                    G10 = book.add_format(property_table_color_shallow)

            # 无最XIAO
            if (ls_skills_other_list[i10] == ls_skills_other_list[i9] and
                    ls_skills_other_list[i9] == ls_skills_other_list[i8] and
                    ls_skills_other_list[i9] == ls_skills_other_list[i7]):
                G1 = book.add_format(property_table_color_blank)
                G2 = book.add_format(property_table_color_blank)
                G3 = book.add_format(property_table_color_blank)
                G4 = book.add_format(property_table_color_blank)
                G5 = book.add_format(property_table_color_blank)
                G6 = book.add_format(property_table_color_blank)
                G7 = book.add_format(property_table_color_blank)
                G8 = book.add_format(property_table_color_blank)
                G9 = book.add_format(property_table_color_blank)
                G10 = book.add_format(property_table_color_blank)



            # 写入单元格
            if ls_skills_other_avg_01 == 0:
                ls_skills_other_avg_01 = '/'
            sheet.merge_range('I180:I181', ls_skills_other_avg_01, G1)
            if ls_skills_other_avg_23 == 0:
                ls_skills_other_avg_23 = '/'
            sheet.merge_range('I182:I183', ls_skills_other_avg_23, G2)
            if ls_skills_other_avg_45 == 0:
                ls_skills_other_avg_45 = '/'
            sheet.merge_range('I184:I185', ls_skills_other_avg_45, G3)
            if ls_skills_other_avg_67 == 0:
                ls_skills_other_avg_67 = '/'
            sheet.merge_range('I186:I187', ls_skills_other_avg_67, G4)
            if ls_skills_other_avg_89 == 0:
                ls_skills_other_avg_89 = '/'
            sheet.merge_range('I188:I189', ls_skills_other_avg_89, G5)
            if ls_skills_other_avg_1011 == 0:
                ls_skills_other_avg_1011 = '/'
            sheet.merge_range('I190:I191', ls_skills_other_avg_1011, G6)
            if ls_skills_other_avg_1213 == 0:
                ls_skills_other_avg_1213 = '/'
            sheet.merge_range('I192:I193', ls_skills_other_avg_1213, G7)
            if ls_skills_other_avg_1415 == 0:
                ls_skills_other_avg_1415 = '/'
            sheet.merge_range('I194:I195', ls_skills_other_avg_1415, G8)
            if ls_skills_other_avg_1617 == 0:
                ls_skills_other_avg_1617 = '/'
            sheet.merge_range('I196:I197', ls_skills_other_avg_1617, G9)
            if ls_skills_other_avg_1819 == 0:
                ls_skills_other_avg_1819 = '/'
            sheet.merge_range('I198:I199', ls_skills_other_avg_1819, G10)

            #
            # 填数（自评均分）
            ls_skills_self_avg_01 = []
            if np.isnan(ls_skills_self_0[i]):
                print('NAN')
            else:
                ls_skills_self_avg_01.append(ls_skills_self_0[i])

            if np.isnan(ls_skills_self_1[i]):
                print('NAN')
            else:
                ls_skills_self_avg_01.append(ls_skills_self_1[i])

            if len(ls_skills_self_avg_01) == 0:
                ls_skills_self_avg_01 = 0
            else:
                ls_skills_self_avg_01 = round(
                    sum(ls_skills_self_avg_01) / len(ls_skills_self_avg_01), 2)
            #
            ls_skills_self_avg_23 = []
            if np.isnan(ls_skills_self_2[i]):
                print('NAN')
            else:
                ls_skills_self_avg_23.append(ls_skills_self_2[i])

            if np.isnan(ls_skills_self_3[i]):
                print('NAN')
            else:
                ls_skills_self_avg_23.append(ls_skills_self_3[i])

            if len(ls_skills_self_avg_23) == 0:
                ls_skills_self_avg_23 = 0
            else:
                ls_skills_self_avg_23 = round(
                    sum(ls_skills_self_avg_23) / len(ls_skills_self_avg_23), 2)
            #
            ls_skills_self_avg_45 = []
            if np.isnan(ls_skills_self_4[i]):
                print('NAN')
            else:
                ls_skills_self_avg_45.append(ls_skills_self_4[i])

            if np.isnan(ls_skills_self_5[i]):
                print('NAN')
            else:
                ls_skills_self_avg_45.append(ls_skills_self_5[i])

            if len(ls_skills_self_avg_45) == 0:
                ls_skills_self_avg_45 = 0
            else:
                ls_skills_self_avg_45 = round(
                    sum(ls_skills_self_avg_45) / len(ls_skills_self_avg_45), 2)
            #
            ls_skills_self_avg_67 = []
            if np.isnan(ls_skills_self_6[i]):
                print('NAN')
            else:
                ls_skills_self_avg_67.append(ls_skills_self_6[i])

            if np.isnan(ls_skills_self_7[i]):
                print('NAN')
            else:
                ls_skills_self_avg_67.append(ls_skills_self_7[i])

            if len(ls_skills_self_avg_67) == 0:
                ls_skills_self_avg_67 = 0
            else:
                ls_skills_self_avg_67 = round(
                    sum(ls_skills_self_avg_67) / len(ls_skills_self_avg_67), 2)
            #
            ls_skills_self_avg_89 = []
            if np.isnan(ls_skills_self_8[i]):
                print('NAN')
            else:
                ls_skills_self_avg_89.append(ls_skills_self_8[i])

            if np.isnan(ls_skills_self_9[i]):
                print('NAN')
            else:
                ls_skills_self_avg_89.append(ls_skills_self_9[i])

            if len(ls_skills_self_avg_89) == 0:
                ls_skills_self_avg_89 = 0
            else:
                ls_skills_self_avg_89 = round(
                    sum(ls_skills_self_avg_89) / len(ls_skills_self_avg_89), 2)

            #
            ls_skills_self_avg_1011 = []
            if np.isnan(ls_skills_self_10[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1011.append(ls_skills_self_10[i])

            if np.isnan(ls_skills_self_11[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1011.append(ls_skills_self_11[i])

            if len(ls_skills_self_avg_1011) == 0:
                ls_skills_self_avg_1011 = 0
            else:
                ls_skills_self_avg_1011 = round(
                    sum(ls_skills_self_avg_1011) / len(ls_skills_self_avg_1011), 2)
            #
            #
            ls_skills_self_avg_1213 = []
            if np.isnan(ls_skills_self_12[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1213.append(ls_skills_self_12[i])

            if np.isnan(ls_skills_self_13[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1213.append(ls_skills_self_13[i])

            if len(ls_skills_self_avg_1213) == 0:
                ls_skills_self_avg_1213 = 0
            else:
                ls_skills_self_avg_1213 = round(
                    sum(ls_skills_self_avg_1213) / len(ls_skills_self_avg_1213), 2)

            #
            ls_skills_self_avg_1415 = []
            if np.isnan(ls_skills_self_14[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1415.append(ls_skills_self_14[i])

            if np.isnan(ls_skills_self_15[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1415.append(ls_skills_self_15[i])

            if len(ls_skills_self_avg_1415) == 0:
                ls_skills_self_avg_1415 = 0
            else:
                ls_skills_self_avg_1415 = round(
                    sum(ls_skills_self_avg_1415) / len(ls_skills_self_avg_1415), 2)
            #
            ls_skills_self_avg_1617 = []
            if np.isnan(ls_skills_self_16[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1617.append(ls_skills_self_16[i])

            if np.isnan(ls_skills_self_17[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1617.append(ls_skills_self_17[i])

            if len(ls_skills_self_avg_1617) == 0:
                ls_skills_self_avg_1617 = 0
            else:
                ls_skills_self_avg_1617 = round(
                    sum(ls_skills_self_avg_1617) / len(ls_skills_self_avg_1617), 2)
            #
            ls_skills_self_avg_1819 = []
            if np.isnan(ls_skills_self_18[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1819.append(ls_skills_self_18[i])

            if np.isnan(ls_skills_self_19[i]):
                print('NAN')
            else:
                ls_skills_self_avg_1819.append(ls_skills_self_19[i])

            if len(ls_skills_self_avg_1819) == 0:
                ls_skills_self_avg_1819 = 0
            else:
                ls_skills_self_avg_1819 = round(
                    sum(ls_skills_self_avg_1819) / len(ls_skills_self_avg_1819), 2)

            # 排序判断(自评分数)
            ls_skills_self_list_1 = [ls_skills_self_avg_01, ls_skills_self_avg_23,
                                     ls_skills_self_avg_45, ls_skills_self_avg_67,
                                     ls_skills_self_avg_89, ls_skills_self_avg_1011,
                                     ls_skills_self_avg_1213, ls_skills_self_avg_1415,
                                     ls_skills_self_avg_1617, ls_skills_self_avg_1819]
            ls_skills_self_list = []
            for si in ls_skills_self_list_1:
                if si == 0:
                    print('NULL')
                else:
                    ls_skills_self_list.append(si)
            print('他评均分：' + str(ls_skills_self_list))
            #
            # format1 = []
            # 从小到大排序
            dic1_sort_self = np.argsort(ls_skills_self_list)  # 用索引引用字典  dict[i]
            dss_len = len(dic1_sort_self)
            # 索引排列
            # 取索引的后四位（最大的4位）
            i4 = dic1_sort_self[dsc_len - 4]  # 倒数第4位
            i3 = dic1_sort_self[dsc_len - 3]  # 倒数第3位
            i2 = dic1_sort_self[dsc_len - 2]  # 倒数第2位
            i1 = dic1_sort_self[dsc_len - 1]  # 倒数第1位

            # 取索引的前四位（最小的4位）
            i7 = dic1_sort_self[3]  # 第4位
            i8 = dic1_sort_self[2]  # 第3位
            i9 = dic1_sort_self[1]  # 第2位
            i10 = dic1_sort_self[0]  # 第1位
            #
            # 判断前3位和后3位的逻辑分开写
            # 先预置单元格格式为不添加任何颜色
            H1 = book.add_format(property_table_color_blank)
            H2 = book.add_format(property_table_color_blank)
            H3 = book.add_format(property_table_color_blank)
            H4 = book.add_format(property_table_color_blank)
            H5 = book.add_format(property_table_color_blank)
            H6 = book.add_format(property_table_color_blank)
            H7 = book.add_format(property_table_color_blank)
            H8 = book.add_format(property_table_color_blank)
            H9 = book.add_format(property_table_color_blank)
            H10 = book.add_format(property_table_color_blank)

            # 判断位数(前3位）（最大值有三个）
            if (ls_skills_self_list[i1] == ls_skills_self_list[i2] and ls_skills_self_list[i2] ==
                ls_skills_self_list[i3] and ls_skills_self_list[i3] != ls_skills_self_list[i4]) or \
                    (ls_skills_self_list[i1] != ls_skills_self_list[i2] and ls_skills_self_list[i2] ==
                     ls_skills_self_list[i3] and ls_skills_self_list[i3] != ls_skills_self_list[i4]) or \
                    (ls_skills_self_list[i1] == ls_skills_self_list[i2] and ls_skills_self_list[i2] !=
                     ls_skills_self_list[i3] and ls_skills_self_list[i3] != ls_skills_self_list[i4]) \
                    or (ls_skills_self_list[i1] != ls_skills_self_list[i2] and ls_skills_self_list[i2] !=
                        ls_skills_self_list[i3] and ls_skills_self_list[i3] != ls_skills_self_list[i4]):
                # 有最大的前3项,定位到最大单元格，然后标出颜色
                # 确定最大值，如何关联到颜色（考虑用键值对的形式）例如找到最大值x，那么其对应的y可以被赋值
                if ls_skills_self_avg_01 == ls_skills_self_list[i1] or ls_skills_self_avg_01 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_01 == ls_skills_self_list[i3]:
                    H1 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_23 == ls_skills_self_list[i1] or ls_skills_self_avg_23 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_23 == ls_skills_self_list[i3]:
                    H2 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_45 == ls_skills_self_list[i1] or ls_skills_self_avg_45 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_45 == ls_skills_self_list[i3]:
                    H3 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_67 == ls_skills_self_list[i1] or ls_skills_self_avg_67 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_67 == ls_skills_self_list[i3]:
                    H4 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_89 == ls_skills_self_list[i1] or ls_skills_self_avg_89 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_89 == ls_skills_self_list[i3]:
                    H5 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1011 == ls_skills_self_list[i1] or ls_skills_self_avg_1011 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_1011 == ls_skills_self_list[i3]:
                    H6 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1213 == ls_skills_self_list[i1] or ls_skills_self_avg_1213 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_1213 == ls_skills_self_list[i3]:
                    H7 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1415 == ls_skills_self_list[i1] or ls_skills_self_avg_1415 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_1415 == ls_skills_self_list[i3]:
                    H8 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1617 == ls_skills_self_list[i1] or ls_skills_self_avg_1617 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_1617 == ls_skills_self_list[i3]:
                    H9 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1819 == ls_skills_self_list[i1] or ls_skills_self_avg_1819 == \
                        ls_skills_self_list[i2] or ls_skills_self_avg_1819 == ls_skills_self_list[i3]:
                    H10 = book.add_format(property_table_color_deep)

            # 最大值有两个（考虑的情况没有穷尽）
            if (ls_skills_self_list[i1] == ls_skills_self_list[i2] and ls_skills_self_list[i3] == \
                ls_skills_self_list[i4]) or (ls_skills_self_list[i1] != ls_skills_self_list[i2] and
                                             ls_skills_self_list[i2] != ls_skills_self_list[i3] and
                                             ls_skills_self_list[i3] == ls_skills_self_list[i4]):
                if ls_skills_self_avg_01 == ls_skills_self_list[i1] or ls_skills_self_avg_01 == \
                        ls_skills_self_list[i2]:
                    H1 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_23 == ls_skills_self_list[i1] or ls_skills_self_avg_23 == \
                        ls_skills_self_list[i2]:
                    H2 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_45 == ls_skills_self_list[i1] or ls_skills_self_avg_45 == \
                        ls_skills_self_list[i2]:
                    H3 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_67 == ls_skills_self_list[i1] or ls_skills_self_avg_67 == \
                        ls_skills_self_list[i2]:
                    H4 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_89 == ls_skills_self_list[i1] or ls_skills_self_avg_89 == \
                        ls_skills_self_list[i2]:
                    H5 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1011 == ls_skills_self_list[i1] or ls_skills_self_avg_1011 == \
                        ls_skills_self_list[i2]:
                    H6 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1213 == ls_skills_self_list[i1] or ls_skills_self_avg_1213 == \
                        ls_skills_self_list[i2]:
                    H7 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1415 == ls_skills_self_list[i1] or ls_skills_self_avg_1415 == \
                        ls_skills_self_list[i2]:
                    H8 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1617 == ls_skills_self_list[i1] or ls_skills_self_avg_1617 == \
                        ls_skills_self_list[i2]:
                    H9 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1819 == ls_skills_self_list[i1] or ls_skills_self_avg_1819 == \
                        ls_skills_self_list[i2]:
                    H10 = book.add_format(property_table_color_deep)

            # 最大值只有一个
            if (ls_skills_self_list[i1] != ls_skills_self_list[i2] and
                    ls_skills_self_list[i2] == ls_skills_self_list[i3] and
                    ls_skills_self_list[i2] == ls_skills_self_list[i4]):
                if ls_skills_self_avg_01 == ls_skills_self_list[i1]:
                    H1 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_23 == ls_skills_self_list[i1]:
                    H2 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_45 == ls_skills_self_list[i1]:
                    H3 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_67 == ls_skills_self_list[i1]:
                    H4 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_89 == ls_skills_self_list[i1]:
                    H5 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1011 == ls_skills_self_list[i1]:
                    H6 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1213 == ls_skills_self_list[i1]:
                    H7 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1415 == ls_skills_self_list[i1]:
                    H8 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1617 == ls_skills_self_list[i1]:
                    H9 = book.add_format(property_table_color_deep)
                if ls_skills_self_avg_1819 == ls_skills_self_list[i1]:
                    H10 = book.add_format(property_table_color_deep)

            # 无最大
            if (ls_skills_self_list[i1] == ls_skills_self_list[i2] and
                    ls_skills_self_list[i2] == ls_skills_self_list[i3] and
                    ls_skills_self_list[i3] == ls_skills_self_list[i4]):
                H1 = book.add_format(property_table_color_blank)
                H2 = book.add_format(property_table_color_blank)
                H3 = book.add_format(property_table_color_blank)
                H4 = book.add_format(property_table_color_blank)
                H5 = book.add_format(property_table_color_blank)
                H6 = book.add_format(property_table_color_blank)
                H7 = book.add_format(property_table_color_blank)
                H8 = book.add_format(property_table_color_blank)
                H9 = book.add_format(property_table_color_blank)
                H10 = book.add_format(property_table_color_blank)

            # 后三位判断（最小值有3位）
            if (ls_skills_self_list[i10] == ls_skills_self_list[i9] and ls_skills_self_list[i9] ==
                ls_skills_self_list[i8] and ls_skills_self_list[i8] != ls_skills_self_list[i7]) or \
                    (ls_skills_self_list[i10] != ls_skills_self_list[i9] and ls_skills_self_list[i9] ==
                     ls_skills_self_list[i8] and
                     ls_skills_self_list[i8] != ls_skills_self_list[i7]) or \
                    (ls_skills_self_list[i10] == ls_skills_self_list[i9] and ls_skills_self_list[i9] !=
                     ls_skills_self_list[i8] and ls_skills_self_list[i8] != ls_skills_self_list[i7]) or \
                    (ls_skills_self_list[i10] != ls_skills_self_list[i9] and ls_skills_self_list[i9] !=
                     ls_skills_self_list[i8] and ls_skills_self_list[i8] != ls_skills_self_list[i7]):
                if ls_skills_self_avg_01 == ls_skills_self_list[i10] or ls_skills_self_avg_01 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_01 == ls_skills_self_list[i8]:
                    H1 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_23 == ls_skills_self_list[i10] or ls_skills_self_avg_23 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_23 == ls_skills_self_list[i8]:
                    H2 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_45 == ls_skills_self_list[i10] or ls_skills_self_avg_45 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_45 == ls_skills_self_list[i8]:
                    H3 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_67 == ls_skills_self_list[i10] or ls_skills_self_avg_67 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_67 == ls_skills_self_list[i8]:
                    H4 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_89 == ls_skills_self_list[i10] or ls_skills_self_avg_89 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_89 == ls_skills_self_list[i8]:
                    H5 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1011 == ls_skills_self_list[i10] or ls_skills_self_avg_1011 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_1011 == ls_skills_self_list[i8]:
                    H6 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1213 == ls_skills_self_list[i10] or ls_skills_self_avg_1213 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_1213 == ls_skills_self_list[i8]:
                    H7 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1415 == ls_skills_self_list[i10] or ls_skills_self_avg_1415 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_1415 == ls_skills_self_list[i8]:
                    H8 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1617 == ls_skills_self_list[i10] or ls_skills_self_avg_1617 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_1617 == ls_skills_self_list[i8]:
                    H9 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1819 == ls_skills_self_list[i10] or ls_skills_self_avg_1819 == \
                        ls_skills_self_list[i9] \
                        or ls_skills_self_avg_1819 == ls_skills_self_list[i8]:
                    H10 = book.add_format(property_table_color_shallow)

            # 最小值有两个（考虑的情况没有穷尽）
            if (ls_skills_self_list[i10] == ls_skills_self_list[i9] and ls_skills_self_list[i8] ==
                ls_skills_self_list[i7]) or \
                    (ls_skills_self_list[i10] != ls_skills_self_list[i9] and ls_skills_self_list[i9] !=
                     ls_skills_self_list[i8] and ls_skills_self_list[i8] == ls_skills_self_list[i7]):
                if ls_skills_self_avg_01 == ls_skills_self_list[i10] or ls_skills_self_avg_01 == \
                        ls_skills_self_list[i9]:
                    H1 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_23 == ls_skills_self_list[i10] or ls_skills_self_avg_23 == \
                        ls_skills_self_list[i9]:
                    H2 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_45 == ls_skills_self_list[i10] or ls_skills_self_avg_45 == \
                        ls_skills_self_list[i9]:
                    H3 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_67 == ls_skills_self_list[i10] or ls_skills_self_avg_67 == \
                        ls_skills_self_list[i9]:
                    H4 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_89 == ls_skills_self_list[i10] or ls_skills_self_avg_89 == \
                        ls_skills_self_list[i9]:
                    H5 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1011 == ls_skills_self_list[i10] or ls_skills_self_avg_1011 == \
                        ls_skills_self_list[i9]:
                    H6 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1213 == ls_skills_self_list[i10] or ls_skills_self_avg_1213 == \
                        ls_skills_self_list[i9]:
                    H7 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1415 == ls_skills_self_list[i10] or ls_skills_self_avg_1415 == \
                        ls_skills_self_list[i9]:
                    H8 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1617 == ls_skills_self_list[i10] or ls_skills_self_avg_1617 == \
                        ls_skills_self_list[i9]:
                    H9 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1819 == ls_skills_self_list[i10] or ls_skills_self_avg_1819 == \
                        ls_skills_self_list[i9]:
                    H10 = book.add_format(property_table_color_shallow)

            # 最小值只有一个
            if (ls_skills_self_list[i10] != ls_skills_self_list[i9] and
                    ls_skills_self_list[i9] == ls_skills_self_list[i8] and
                    ls_skills_self_list[i9] == ls_skills_self_list[i7]):
                if ls_skills_self_avg_01 == ls_skills_self_list[i10]:
                    H1 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_23 == ls_skills_self_list[i10]:
                    H2 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_45 == ls_skills_self_list[i10]:
                    H3 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_67 == ls_skills_self_list[i10]:
                    H4 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_89 == ls_skills_self_list[i10]:
                    H5 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1011 == ls_skills_self_list[i10]:
                    H6 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1213 == ls_skills_self_list[i10]:
                    H7 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1415 == ls_skills_self_list[i10]:
                    H8 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1617 == ls_skills_self_list[i10]:
                    H9 = book.add_format(property_table_color_shallow)
                if ls_skills_self_avg_1819 == ls_skills_self_list[i10]:
                    H10 = book.add_format(property_table_color_shallow)

            # 无最xiao
            if (ls_skills_self_list[i10] == ls_skills_self_list[i9] and
                    ls_skills_self_list[i9] == ls_skills_self_list[i8] and
                    ls_skills_self_list[i8] == ls_skills_self_list[i7]):
                H1 = book.add_format(property_table_color_blank)
                H2 = book.add_format(property_table_color_blank)
                H3 = book.add_format(property_table_color_blank)
                H4 = book.add_format(property_table_color_blank)
                H5 = book.add_format(property_table_color_blank)
                H6 = book.add_format(property_table_color_blank)
                H7 = book.add_format(property_table_color_blank)
                H8 = book.add_format(property_table_color_blank)
                H9 = book.add_format(property_table_color_blank)
                H10 = book.add_format(property_table_color_blank)



            # # 自评
            if ls_skills_self_avg_01 == 0:
                ls_skills_self_avg_01 = '/'
            sheet.merge_range('J180:J181', ls_skills_self_avg_01, H1)
            if ls_skills_self_avg_23 == 0:
                ls_skills_self_avg_23 = '/'
            sheet.merge_range('J182:J183', ls_skills_self_avg_23, H2)
            if ls_skills_self_avg_45 == 0:
                ls_skills_self_avg_45 = '/'
            sheet.merge_range('J184:J185', ls_skills_self_avg_45, H3)
            if ls_skills_self_avg_67 == 0:
                ls_skills_self_avg_67 = '/'
            sheet.merge_range('J186:J187', ls_skills_self_avg_67, H4)
            if ls_skills_self_avg_89 == 0:
                ls_skills_self_avg_89 = '/'
            sheet.merge_range('J188:J189', ls_skills_self_avg_89, H5)
            if ls_skills_self_avg_1011 == 0:
                ls_skills_self_avg_1011 = '/'
            sheet.merge_range('J190:J191', ls_skills_self_avg_1011, H6)
            if ls_skills_self_avg_1213 == 0:
                ls_skills_self_avg_1213 = '/'
            sheet.merge_range('J192:J193', ls_skills_self_avg_1213, H7)
            if ls_skills_self_avg_1415 == 0:
                ls_skills_self_avg_1415 = '/'
            sheet.merge_range('J194:J195', ls_skills_self_avg_1415, H8)
            if ls_skills_self_avg_1617 == 0:
                ls_skills_self_avg_1617 = '/'
            sheet.merge_range('J196:J197', ls_skills_self_avg_1617, H9)
            if ls_skills_self_avg_1819 == 0:
                ls_skills_self_avg_1819 = '/'
            sheet.merge_range('J198:J199', ls_skills_self_avg_1819, H10)
            # #

            #  在MO的领导技能中，你他评得分最高的行为包括：
            sheet.merge_range('B201:F201', "在MO的领导技能中，你的“他评均分”最高的行为包括：", cell_format_content)

            # 排序判断(下级评分)
            # 能力项
            skill0 = '关注市场、客户和本领域动态'
            skill1 = '理解公司方向和部门策略重点'
            skill2 = '基于数据和事实分析、解决问题'
            skill3 = '识别高质量人才，帮助新人快速融入'
            skill4 = '及时提供反馈、分享个人经验，帮助他人成长'
            skill5 = '识通过奖励和认可形成一个积极向上的团队氛围'
            skill6 = '通过有效分工、计划、跟进等确保高效执行'
            skill7 = '有行动力，能带领团队保质、保量、如期地完成任务'
            skill8 = '切实通过复盘总结经验教训，不断改进'
            skill9 = '通过主动找不足和差距不断提升自己'
            # 行为项
            active0 = '能定期收集并整合客户与市场的信息'
            active1 = '可以持续关注本专业&领域的最新动态，为团队提供输入'
            active2 = '及时了解公司最新的发展方向与战略重点，并给团队分享'
            active3 = '能根据部门的策略重点有序的安排与推进工作'
            active4 = '根据数据与事实有逻辑地分析问题，例如部门的目标设定标准与业务发展方向'
            active5 = '对部门的问题能提出专业有效的决策建议'
            active6 = '善于观察与识别不同人的特质与长短板，甄别与吸引优秀的人才'
            active7 = '营造团队互信与友善的协作氛围，帮助新人快速融入团队'
            active8 = '基于下属在工作中的行为给出及时与具体的激励型反馈与建设型反馈'
            active9 = '用合理的方式分享自身经验，辅导员工提升完成任务的能力'
            active10 = '定期与下属沟通，并在过程中能持续帮助下属理解公司价值观'
            active11 = '能有效识别不同下属的激励点，并给出有针对性及有效的激励方式'
            active12 = '根据团队成员的特点合理分配任务'
            active13 = '与下属共识详细任务实施计划，并及时跟踪检查进展情况'
            active14 = '带领团队快速行动排除影响任务完成的障碍'
            active15 = '必要时及时调整计划和人员配置确保任务目标达成'
            active16 = '引导团队通过规范的复盘方式不断总结经验与教训，养成复盘的习惯'
            active17 = '带领团队积极寻找持续优化工作方法与工作流程的机会'
            active18 = '主动寻求并接纳他人的意见和反馈，不断自我改进'
            active19 = '可以持续探索和学习新的知识和领域'
            #
            skill_active_list = []
            skill_active_list_l5 = []
            skill_active_list_1 = []
            skill_active_list_2 = []
            # loc1_s_a = 202
            # loc1_s_b = 202
            #
            s_1 = '能力【'
            s_2 = '】中的'
            a_1 = '行为【'
            a_2 = '】'
            #
            #

            ls_skills_other_list_f1 = [ls_skills_other_0[i], ls_skills_other_1[i], ls_skills_other_2[i],
                                       ls_skills_other_3[i], ls_skills_other_4[i],
                                       ls_skills_other_5[i], ls_skills_other_6[i], ls_skills_other_7[i],
                                       ls_skills_other_8[i], ls_skills_other_9[i],
                                       ls_skills_other_10[i], ls_skills_other_11[i], ls_skills_other_12[i],
                                       ls_skills_other_13[i],
                                       ls_skills_other_14[i], ls_skills_other_15[i], ls_skills_other_16[i],
                                       ls_skills_other_17[i],
                                       ls_skills_other_18[i], ls_skills_other_19[i]]
            ls_skills_other_list_f = []
            for isofi in ls_skills_other_list_f1:
                if np.isnan(isofi):
                    print('NULL')
                else:
                    ls_skills_other_list_f.append(isofi)
            print('他评均分：' + str(ls_skills_other_list_f))
            # print(ls_skills_other_list_f)
            #
            # 从小到大排序
            dic1_sort_ls_skills_other_list_f = np.argsort(ls_skills_other_list_f)
            dsisolf_len = len(dic1_sort_ls_skills_other_list_f)
            # print(dic1_sort_ls_skills_other_list_f)
            # 索引排列
            # 取索引的后四位（最大的5位）
            i6 = dic1_sort_ls_skills_other_list_f[dsisolf_len - 6]  # 第6位
            i5 = dic1_sort_ls_skills_other_list_f[dsisolf_len - 5]  # 第5位
            i4 = dic1_sort_ls_skills_other_list_f[dsisolf_len - 4]  # 第4位
            i3 = dic1_sort_ls_skills_other_list_f[dsisolf_len - 3]  # 第3位
            i2 = dic1_sort_ls_skills_other_list_f[dsisolf_len - 2]  # 第2位
            i1 = dic1_sort_ls_skills_other_list_f[dsisolf_len - 1]  # 第1位

            # 取索引的前四位（最小的5位）
            i14 = dic1_sort_ls_skills_other_list_f[5]  # 第4位
            i15 = dic1_sort_ls_skills_other_list_f[4]  # 第5位
            i16 = dic1_sort_ls_skills_other_list_f[3]  # 第4位
            i17 = dic1_sort_ls_skills_other_list_f[2]  # 第3位
            i18 = dic1_sort_ls_skills_other_list_f[1]  # 第2位
            i19 = dic1_sort_ls_skills_other_list_f[0]  # 第1位

            #
            print(ls_skills_other_list_f[i19], ls_skills_other_list_f[i18], ls_skills_other_list_f[i17],
                  ls_skills_other_list_f[i16]
                  , ls_skills_other_list_f[i15], ls_skills_other_list_f[i14])
            print(ls_skills_other_list_f[i1], ls_skills_other_list_f[i2], ls_skills_other_list_f[i3],
                  ls_skills_other_list_f[i4]
                  , ls_skills_other_list_f[i5], ls_skills_other_list_f[i6])

            # m = [ls_skills_other_list_f[i1], ls_skills_other_list_f[i2], ls_skills_other_list_f[i3],
            #      ls_skills_other_list_f[i4], ls_skills_other_list_f[i5],
            #      ls_skills_other_list_f[i6], ls_skills_other_list_f[i14], ls_skills_other_list_f[i15],
            #      ls_skills_other_list_f[i16], ls_skills_other_list_f[i17],
            #      ls_skills_other_list_f[i18], ls_skills_other_list_f[i19]]
            # print(m)
            # 判断位数(前5位）（最大值有5个） 10个
            if (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] !=
                     ls_skills_other_list_f[i6]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i1] or ls_skills_other_0[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_0[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_0[i] == ls_skills_other_list_f[i4] or ls_skills_other_0[i] == \
                        ls_skills_other_list_f[i5]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i1] or ls_skills_other_1[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_1[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_1[i] == ls_skills_other_list_f[i4] or ls_skills_other_1[i] == \
                        ls_skills_other_list_f[i5]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i1] or ls_skills_other_2[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_2[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_2[i] == ls_skills_other_list_f[i4] or ls_skills_other_2[i] == \
                        ls_skills_other_list_f[i5]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i1] or ls_skills_other_3[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_3[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_3[i] == ls_skills_other_list_f[i4] or ls_skills_other_3[i] == \
                        ls_skills_other_list_f[i5]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i1] or ls_skills_other_4[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_4[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_4[i] == ls_skills_other_list_f[i4] or ls_skills_other_4[i] == \
                        ls_skills_other_list_f[i5]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i1] or ls_skills_other_5[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_5[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_5[i] == ls_skills_other_list_f[i4] or ls_skills_other_5[i] == \
                        ls_skills_other_list_f[i5]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i1] or ls_skills_other_6[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_6[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_6[i] == ls_skills_other_list_f[i4] or ls_skills_other_6[i] == \
                        ls_skills_other_list_f[i5]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i1] or ls_skills_other_7[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_7[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_7[i] == ls_skills_other_list_f[i4] or ls_skills_other_7[i] == \
                        ls_skills_other_list_f[i5]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i1] or ls_skills_other_8[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_8[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_8[i] == ls_skills_other_list_f[i4] or ls_skills_other_8[i] == \
                        ls_skills_other_list_f[i5]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i1] or ls_skills_other_9[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_9[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_9[i] == ls_skills_other_list_f[i4] or ls_skills_other_9[i] == \
                        ls_skills_other_list_f[i5]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i1] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_10[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_10[i] == ls_skills_other_list_f[i4] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[i5]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i1] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_11[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_11[i] == ls_skills_other_list_f[i4] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[i5]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i1] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_12[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_12[i] == ls_skills_other_list_f[i4] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[i5]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i1] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_13[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_13[i] == ls_skills_other_list_f[i4] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[i5]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i1] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_14[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_14[i] == ls_skills_other_list_f[i4] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[i5]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i1] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_15[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_15[i] == ls_skills_other_list_f[i4] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[i5]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i1] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_16[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_16[i] == ls_skills_other_list_f[i4] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[i5]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i1] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_17[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_17[i] == ls_skills_other_list_f[i4] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[i5]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i1] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_18[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_18[i] == ls_skills_other_list_f[i4] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[i5]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i1] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_19[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_19[i] == ls_skills_other_list_f[i4] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[i5]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list.append(str19)
                print('他评列表:' + str(skill_active_list))
                for item in skill_active_list:  # 5个
                    print(item)
                    if len(item) != 0:
                        skill_active_list_1.append(item)
                skill_active_list_1 = "\n".join(skill_active_list_1)  # 换行符连接
                sheet.merge_range('B202:J206', skill_active_list_1, cell_format_content_skill)

            # 4个(待核查)
            elif (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                  ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                  ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                  ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                     ls_skills_other_list_f[i6]) or (
                    ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                    ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                    ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                    ls_skills_other_list_f[i6]) or (
                    ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                    ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                    ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                    ls_skills_other_list_f[i6]) or (
                    ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                    ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                    ls_skills_other_list_f[i4] != ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                    ls_skills_other_list_f[i6]):

                if ls_skills_other_0[i] == ls_skills_other_list_f[i1] or ls_skills_other_0[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_0[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_0[i] == ls_skills_other_list_f[i4]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i1] or ls_skills_other_1[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_1[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_1[i] == ls_skills_other_list_f[i4]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i1] or ls_skills_other_2[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_2[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_2[i] == ls_skills_other_list_f[i4]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i1] or ls_skills_other_3[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_3[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_3[i] == ls_skills_other_list_f[i4]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i1] or ls_skills_other_4[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_4[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_4[i] == ls_skills_other_list_f[i4]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i1] or ls_skills_other_5[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_5[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_5[i] == ls_skills_other_list_f[i4]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i1] or ls_skills_other_6[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_6[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_6[i] == ls_skills_other_list_f[i4]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i1] or ls_skills_other_7[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_7[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_7[i] == ls_skills_other_list_f[i4]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i1] or ls_skills_other_8[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_8[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_8[i] == ls_skills_other_list_f[i4]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i1] or ls_skills_other_9[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_9[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_9[i] == ls_skills_other_list_f[i4]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i1] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_10[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_10[i] == ls_skills_other_list_f[i4]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i1] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_11[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_11[i] == ls_skills_other_list_f[i4]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i1] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_12[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_12[i] == ls_skills_other_list_f[i4]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i1] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_13[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_13[i] == ls_skills_other_list_f[i4]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i1] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_14[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_14[i] == ls_skills_other_list_f[i4]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i1] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_15[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_15[i] == ls_skills_other_list_f[i4]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i1] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_16[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_16[i] == ls_skills_other_list_f[i4]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i1] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_17[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_17[i] == ls_skills_other_list_f[i4]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i1] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_18[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_18[i] == ls_skills_other_list_f[i4]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i1] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_19[i] == ls_skills_other_list_f[i3] or \
                        ls_skills_other_19[i] == ls_skills_other_list_f[i4]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list.append(str19)
                print(len(skill_active_list))
                for item in skill_active_list:  # 4个
                    print(item)
                    if len(item) != 0:
                        skill_active_list_1.append(item)
                skill_active_list_1 = "\n".join(skill_active_list_1)  # 换行符连接
                print(skill_active_list_1)
                sheet.merge_range('B202:J205', skill_active_list_1, cell_format_content_skill)

            # 3个
            elif (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                  ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                  ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                  ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                     ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] != ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                     ls_skills_other_list_f[i6]):

                if ls_skills_other_0[i] == ls_skills_other_list_f[i1] or ls_skills_other_0[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_0[i] == ls_skills_other_list_f[i3]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i1] or ls_skills_other_1[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_1[i] == ls_skills_other_list_f[i3]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i1] or ls_skills_other_2[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_2[i] == ls_skills_other_list_f[i3]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i1] or ls_skills_other_3[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_3[i] == ls_skills_other_list_f[i3]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i1] or ls_skills_other_4[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_4[i] == ls_skills_other_list_f[i3]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i1] or ls_skills_other_5[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_5[i] == ls_skills_other_list_f[i3]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i1] or ls_skills_other_6[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_6[i] == ls_skills_other_list_f[i3]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i1] or ls_skills_other_7[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_7[i] == ls_skills_other_list_f[i3]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i1] or ls_skills_other_8[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_8[i] == ls_skills_other_list_f[i3]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i1] or ls_skills_other_9[i] == ls_skills_other_list_f[
                    i2] or ls_skills_other_9[i] == ls_skills_other_list_f[i3]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i1] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_10[i] == ls_skills_other_list_f[i3]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i1] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_11[i] == ls_skills_other_list_f[i3]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i1] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_12[i] == ls_skills_other_list_f[i3]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i1] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_13[i] == ls_skills_other_list_f[i3]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i1] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_14[i] == ls_skills_other_list_f[i3]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i1] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_15[i] == ls_skills_other_list_f[i3]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i1] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_16[i] == ls_skills_other_list_f[i3]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i1] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_17[i] == ls_skills_other_list_f[i3]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i1] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_18[i] == ls_skills_other_list_f[i3]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i1] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[i2] or ls_skills_other_19[i] == ls_skills_other_list_f[i3]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list.append(str19)
                print(len(skill_active_list))
                for item in skill_active_list:  # 4个
                    if len(item) != 0:
                        skill_active_list_1.append(item)
                skill_active_list_1 = "\n".join(skill_active_list_1)  # 换行符连接
                print(skill_active_list_1)
                sheet.merge_range('B202:J204', skill_active_list_1, cell_format_content_skill)

            # 2个
            elif (ls_skills_other_list_f[i1] == ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                  ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                  ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                  ls_skills_other_list_f[i6]) or \
                    (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] !=
                     ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                     ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                     ls_skills_other_list_f[i6]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i1] or ls_skills_other_0[i] == ls_skills_other_list_f[
                    i2]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i1] or ls_skills_other_1[i] == ls_skills_other_list_f[
                    i2]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i1] or ls_skills_other_2[i] == ls_skills_other_list_f[
                    i2]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i1] or ls_skills_other_3[i] == ls_skills_other_list_f[
                    i2]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i1] or ls_skills_other_4[i] == ls_skills_other_list_f[
                    i2]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i1] or ls_skills_other_5[i] == ls_skills_other_list_f[
                    i2]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i1] or ls_skills_other_6[i] == ls_skills_other_list_f[
                    i2]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i1] or ls_skills_other_7[i] == ls_skills_other_list_f[
                    i2]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i1] or ls_skills_other_8[i] == ls_skills_other_list_f[
                    i2]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i1] or ls_skills_other_9[i] == ls_skills_other_list_f[
                    i2]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i1] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[i2]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i1] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[i2]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i1] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[i2]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i1] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[i2]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i1] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[i2]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i1] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[i2]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i1] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[i2]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i1] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[i2]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i1] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[i2]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i1] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[i2]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list.append(str19)
                print(len(skill_active_list))
                for item in skill_active_list:  # 4个
                    if len(item) != 0:
                        skill_active_list_1.append(item)
                skill_active_list_1 = "\n".join(skill_active_list_1)  # 换行符连接
                print(skill_active_list_1)
                sheet.merge_range('B202:J203', skill_active_list_1, cell_format_content_skill)

            # 1个
            elif (ls_skills_other_list_f[i1] != ls_skills_other_list_f[i2] and ls_skills_other_list_f[i2] ==
                  ls_skills_other_list_f[i3] and ls_skills_other_list_f[i3] == ls_skills_other_list_f[i4] and
                  ls_skills_other_list_f[i4] == ls_skills_other_list_f[i5] and ls_skills_other_list_f[i5] ==
                  ls_skills_other_list_f[i6]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i1]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i1]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i1]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i1]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i1]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i1]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i1]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i1]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i1]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i1]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i1]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i1]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i1]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i1]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i1]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i1]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i1]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i1]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i1]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i1]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list.append(str19)
                print(len(skill_active_list))
                for item in skill_active_list:  # 4个
                    if len(item) != 0:
                        skill_active_list_1.append(item)
                skill_active_list_1 = "\n".join(skill_active_list_1)  # 换行符连接
                print(skill_active_list_1)
                sheet.merge_range('B202:J202', skill_active_list_1, cell_format_content_skill)

            # ”显示文字“由于你的行为得分同分数较多，此处不做显示“
            else:
                sheet.merge_range('B202:J202', '由于你的行为得分同分数较多，此处不做显示', cell_format_content_skill)

            # 根据最高他评得分的条数，判断接下来单元格的位置
            q = 'B'
            w = ':J'
            w1 = ':M'
            e = 202
            y = 206

            if len(skill_active_list) == 5:
                d = e + 7
            elif len(skill_active_list) == 4:
                d = e + 6
            elif len(skill_active_list) == 3:
                d = e + 5
            elif len(skill_active_list) == 2:
                d = e + 4
            elif len(skill_active_list) == 1:
                d = e + 3
            else:
                d = e + 2
                # loc_other_skill = q + str(e+1) + w + str(e+1)
                # loc_other_skill = 'B203:K203'

            # 在MO的领导技能中，你他评得分最低的行为包括：
            print('高分列表长度：' + str(len(skill_active_list)))
            loc_other_skill = q + str(d - 2) + w + str(d - 2)
            print(loc_other_skill)
            sheet.merge_range(loc_other_skill, "在MO的领导技能中，你的”他评均分“最低的行为包括：", cell_format_content)
            #

            #
            # 判断位数(前5位）（最小值有5个）
            if (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] !=
                     ls_skills_other_list_f[i14]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i19] or ls_skills_other_0[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_0[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_0[i] == ls_skills_other_list_f[i16] or ls_skills_other_0[i] == \
                        ls_skills_other_list_f[i15]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list_l5.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i19] or ls_skills_other_1[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_1[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_1[i] == ls_skills_other_list_f[i16] or ls_skills_other_1[i] == \
                        ls_skills_other_list_f[i15]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list_l5.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i19] or ls_skills_other_2[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_2[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_2[i] == ls_skills_other_list_f[i16] or ls_skills_other_2[i] == \
                        ls_skills_other_list_f[i15]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list_l5.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i19] or ls_skills_other_3[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_3[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_3[i] == ls_skills_other_list_f[i16] or ls_skills_other_3[i] == \
                        ls_skills_other_list_f[i15]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list_l5.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i19] or ls_skills_other_4[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_4[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_4[i] == ls_skills_other_list_f[i16] or ls_skills_other_4[i] == \
                        ls_skills_other_list_f[i15]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list_l5.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i19] or ls_skills_other_5[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_5[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_5[i] == ls_skills_other_list_f[i16] or ls_skills_other_5[i] == \
                        ls_skills_other_list_f[i15]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list_l5.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i19] or ls_skills_other_6[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_6[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_6[i] == ls_skills_other_list_f[i16] or ls_skills_other_6[i] == \
                        ls_skills_other_list_f[i15]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list_l5.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i19] or ls_skills_other_7[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_7[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_7[i] == ls_skills_other_list_f[i16] or ls_skills_other_7[i] == \
                        ls_skills_other_list_f[i15]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list_l5.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i19] or ls_skills_other_8[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_8[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_8[i] == ls_skills_other_list_f[i16] or ls_skills_other_8[i] == \
                        ls_skills_other_list_f[i15]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list_l5.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i19] or ls_skills_other_9[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_9[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_9[i] == ls_skills_other_list_f[i16] or ls_skills_other_9[i] == \
                        ls_skills_other_list_f[i15]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list_l5.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i19] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_10[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_10[i] == ls_skills_other_list_f[i16] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[i15]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list_l5.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i19] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_11[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_11[i] == ls_skills_other_list_f[i16] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[i15]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list_l5.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i19] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_12[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_12[i] == ls_skills_other_list_f[i16] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[i15]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list_l5.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i19] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_13[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_13[i] == ls_skills_other_list_f[i16] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[i15]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list_l5.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i19] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_14[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_14[i] == ls_skills_other_list_f[i16] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[i15]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list_l5.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i19] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_15[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_15[i] == ls_skills_other_list_f[i16] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[i15]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list_l5.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i19] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_16[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_16[i] == ls_skills_other_list_f[i16] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[i15]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list_l5.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i19] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_17[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_17[i] == ls_skills_other_list_f[i16] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[i15]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list_l5.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i19] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_18[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_18[i] == ls_skills_other_list_f[i16] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[i15]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list_l5.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i19] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_19[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_19[i] == ls_skills_other_list_f[i16] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[i15]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list_l5.append(str19)

                if len(skill_active_list) == 5:
                    r = e + 7
                    o = y + 7
                elif len(skill_active_list) == 4:
                    r = e + 6
                    o = y + 6
                elif len(skill_active_list) == 3:
                    r = e + 5
                    o = y + 5
                elif len(skill_active_list) == 2:
                    r = e + 4
                    o = y + 4
                elif len(skill_active_list) == 1:
                    r = e + 3
                    o = y + 3
                else:
                    r = e + 2
                    o = y + 2

                for item_1 in skill_active_list_l5:  # 5个
                    if len(item_1) != 0:
                        skill_active_list_2.append(item_1)
                skill_active_list_2 = "\n".join(skill_active_list_2)  # 换行符连接
                print('低分列表长度：' + str(len(skill_active_list_l5)))
                print(skill_active_list_2)
                r_1 = q + str(r - 1) + w + str(r + len(skill_active_list_l5) - 2)
                print('最小值的位置' + r_1)
                sheet.merge_range(r_1, skill_active_list_2, cell_format_content_skill)

            # 4个(待核查)：最小值
            elif (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                  ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                  ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                  ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                     ls_skills_other_list_f[i14]) or (
                    ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                    ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                    ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                    ls_skills_other_list_f[i14]) or (
                    ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                    ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                    ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                    ls_skills_other_list_f[i14]) or (
                    ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                    ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                    ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                    ls_skills_other_list_f[i14]) or (
                    ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                    ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                    ls_skills_other_list_f[i16] != ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                    ls_skills_other_list_f[i14]):

                if ls_skills_other_0[i] == ls_skills_other_list_f[i19] or ls_skills_other_0[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_0[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_0[i] == ls_skills_other_list_f[i16]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list_l5.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i19] or ls_skills_other_1[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_1[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_1[i] == ls_skills_other_list_f[i16]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list_l5.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i19] or ls_skills_other_2[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_2[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_2[i] == ls_skills_other_list_f[i16]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list_l5.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i19] or ls_skills_other_3[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_3[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_3[i] == ls_skills_other_list_f[i16]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list_l5.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i19] or ls_skills_other_4[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_4[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_4[i] == ls_skills_other_list_f[i16]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list_l5.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i19] or ls_skills_other_5[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_5[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_5[i] == ls_skills_other_list_f[i16]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list_l5.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i19] or ls_skills_other_6[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_6[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_6[i] == ls_skills_other_list_f[i16]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list_l5.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i19] or ls_skills_other_7[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_7[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_7[i] == ls_skills_other_list_f[i16]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list_l5.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i19] or ls_skills_other_8[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_8[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_8[i] == ls_skills_other_list_f[i16]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list_l5.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i19] or ls_skills_other_9[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_9[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_9[i] == ls_skills_other_list_f[i16]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list_l5.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i19] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_10[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_10[i] == ls_skills_other_list_f[i16]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list_l5.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i19] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_11[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_11[i] == ls_skills_other_list_f[i16]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list_l5.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i19] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_12[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_12[i] == ls_skills_other_list_f[i16]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list_l5.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i19] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_13[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_13[i] == ls_skills_other_list_f[i16]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list_l5.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i19] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_14[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_14[i] == ls_skills_other_list_f[i16]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list_l5.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i19] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_15[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_15[i] == ls_skills_other_list_f[i16]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list_l5.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i19] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_16[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_16[i] == ls_skills_other_list_f[i16]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list_l5.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i19] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_17[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_17[i] == ls_skills_other_list_f[i16]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list_l5.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i19] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_18[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_18[i] == ls_skills_other_list_f[i16]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list_l5.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i19] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_19[i] == ls_skills_other_list_f[i17] or \
                        ls_skills_other_19[i] == ls_skills_other_list_f[i16]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list_l5.append(str19)

                if len(skill_active_list) == 5:
                    r = e + 7
                    o = y + 7
                elif len(skill_active_list) == 4:
                    r = e + 6
                    o = y + 6
                elif len(skill_active_list) == 3:
                    r = e + 5
                    o = y + 5
                elif len(skill_active_list) == 2:
                    r = e + 4
                    o = y + 4
                elif len(skill_active_list) == 1:
                    r = e + 3
                    o = y + 3
                else:
                    r = e + 2
                    o = y + 2

                for item_1 in skill_active_list_l5:  # 5个
                    if len(item_1) != 0:
                        skill_active_list_2.append(item_1)
                skill_active_list_2 = "\n".join(skill_active_list_2)  # 换行符连接
                print('低分列表长度：' + str(len(skill_active_list_l5)))
                print(skill_active_list_2)
                # r_1 = q + str(r) + w + str(o)
                r_1 = q + str(r - 1) + w + str(r + len(skill_active_list_l5) - 2)
                print('最小值的位置' + r_1)
                sheet.merge_range(r_1, skill_active_list_2, cell_format_content_skill)

            # 3个
            elif (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                  ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                  ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                  ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                     ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] != ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                     ls_skills_other_list_f[i14]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i19] or ls_skills_other_0[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_0[i] == ls_skills_other_list_f[i17]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list_l5.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i19] or ls_skills_other_1[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_1[i] == ls_skills_other_list_f[i17]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list_l5.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i19] or ls_skills_other_2[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_2[i] == ls_skills_other_list_f[i17]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list_l5.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i19] or ls_skills_other_3[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_3[i] == ls_skills_other_list_f[i17]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list_l5.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i19] or ls_skills_other_4[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_4[i] == ls_skills_other_list_f[i17]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list_l5.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i19] or ls_skills_other_5[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_5[i] == ls_skills_other_list_f[i17]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list_l5.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i19] or ls_skills_other_6[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_6[i] == ls_skills_other_list_f[i17]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list_l5.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i19] or ls_skills_other_7[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_7[i] == ls_skills_other_list_f[i17]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list_l5.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i19] or ls_skills_other_8[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_8[i] == ls_skills_other_list_f[i17]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list_l5.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i19] or ls_skills_other_9[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_9[i] == ls_skills_other_list_f[i17]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list_l5.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i19] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_10[i] == ls_skills_other_list_f[i17]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list_l5.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i19] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_11[i] == ls_skills_other_list_f[i17]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list_l5.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i19] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_12[i] == ls_skills_other_list_f[i17]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list_l5.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i19] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_13[i] == ls_skills_other_list_f[i17]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list_l5.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i19] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_14[i] == ls_skills_other_list_f[i17]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list_l5.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i19] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_15[i] == ls_skills_other_list_f[i17]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list_l5.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i19] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_16[i] == ls_skills_other_list_f[i17]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list_l5.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i19] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_17[i] == ls_skills_other_list_f[i17]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list_l5.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i19] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_18[i] == ls_skills_other_list_f[i17]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list_l5.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i19] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[
                            i18] or ls_skills_other_19[i] == ls_skills_other_list_f[i17]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list_l5.append(str19)

                if len(skill_active_list) == 5:
                    r = e + 7
                    o = y + 7
                elif len(skill_active_list) == 4:
                    r = e + 6
                    o = y + 6
                elif len(skill_active_list) == 3:
                    r = e + 5
                    o = y + 5
                elif len(skill_active_list) == 2:
                    r = e + 4
                    o = y + 4
                elif len(skill_active_list) == 1:
                    r = e + 3
                    o = y + 3
                else:
                    r = e + 2
                    o = y + 2

                for item_1 in skill_active_list_l5:  # 5个
                    if len(item_1) != 0:
                        skill_active_list_2.append(item_1)
                skill_active_list_2 = "\n".join(skill_active_list_2)  # 换行符连接
                print(skill_active_list_2)
                print('低分列表长度：' + str(len(skill_active_list_l5)))
                # r_1 = q + str(r) + w + str(o)
                r_1 = q + str(r - 1) + w + str(r + len(skill_active_list_l5) - 2)
                print(r_1)
                sheet.merge_range(r_1, skill_active_list_2, cell_format_content_skill)

            # 2个
            elif (ls_skills_other_list_f[i19] == ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                  ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                  ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                  ls_skills_other_list_f[i14]) or \
                    (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] !=
                     ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                     ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                     ls_skills_other_list_f[i14]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i19] or ls_skills_other_0[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list_l5.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i19] or ls_skills_other_1[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list_l5.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i19] or ls_skills_other_2[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list_l5.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i19] or ls_skills_other_3[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list_l5.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i19] or ls_skills_other_4[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list_l5.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i19] or ls_skills_other_5[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list_l5.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i19] or ls_skills_other_6[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list_l5.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i19] or ls_skills_other_7[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list_l5.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i19] or ls_skills_other_8[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list_l5.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i19] or ls_skills_other_9[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list_l5.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i19] or ls_skills_other_10[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list_l5.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i19] or ls_skills_other_11[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list_l5.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i19] or ls_skills_other_12[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list_l5.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i19] or ls_skills_other_13[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list_l5.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i19] or ls_skills_other_14[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list_l5.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i19] or ls_skills_other_15[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list_l5.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i19] or ls_skills_other_16[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list_l5.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i19] or ls_skills_other_17[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list_l5.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i19] or ls_skills_other_18[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list_l5.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i19] or ls_skills_other_19[i] == \
                        ls_skills_other_list_f[
                            i18]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list_l5.append(str19)

                if len(skill_active_list) == 5:
                    r = e + 7
                    o = y + 7
                elif len(skill_active_list) == 4:
                    r = e + 6
                    o = y + 6
                elif len(skill_active_list) == 3:
                    r = e + 5
                    o = y + 5
                elif len(skill_active_list) == 2:
                    r = e + 4
                    o = y + 4
                elif len(skill_active_list) == 1:
                    r = e + 3
                    o = y + 3
                else:
                    r = e + 2
                    o = y + 2

                for item_1 in skill_active_list_l5:  # 5个
                    if len(item_1) != 0:
                        skill_active_list_2.append(item_1)
                skill_active_list_2 = "\n".join(skill_active_list_2)  # 换行符连接
                print(skill_active_list_2)
                print('低分列表长度：' + str(len(skill_active_list_l5)))
                # r_1 = q + str(r) + w + str
                r_1 = q + str(r - 1) + w + str(r + len(skill_active_list_l5) - 2)
                print(r_1)
                sheet.merge_range(r_1, skill_active_list_2, cell_format_content_skill)

            # 1个
            elif (ls_skills_other_list_f[i19] != ls_skills_other_list_f[i18] and ls_skills_other_list_f[i18] ==
                  ls_skills_other_list_f[i17] and ls_skills_other_list_f[i17] == ls_skills_other_list_f[i16] and
                  ls_skills_other_list_f[i16] == ls_skills_other_list_f[i15] and ls_skills_other_list_f[i15] ==
                  ls_skills_other_list_f[i14]):
                if ls_skills_other_0[i] == ls_skills_other_list_f[i19]:
                    str0 = s_1 + skill0 + s_2 + a_1 + active0 + a_2
                    skill_active_list_l5.append(str0)
                if ls_skills_other_1[i] == ls_skills_other_list_f[i19]:
                    str1 = s_1 + skill0 + s_2 + a_1 + active1 + a_2
                    skill_active_list_l5.append(str1)
                if ls_skills_other_2[i] == ls_skills_other_list_f[i19]:
                    str2 = s_1 + skill1 + s_2 + a_1 + active2 + a_2
                    skill_active_list_l5.append(str2)
                if ls_skills_other_3[i] == ls_skills_other_list_f[i19]:
                    str3 = s_1 + skill1 + s_2 + a_1 + active3 + a_2
                    skill_active_list_l5.append(str3)
                if ls_skills_other_4[i] == ls_skills_other_list_f[i19]:
                    str4 = s_1 + skill2 + s_2 + a_1 + active4 + a_2
                    skill_active_list_l5.append(str4)
                if ls_skills_other_5[i] == ls_skills_other_list_f[i19]:
                    str5 = s_1 + skill2 + s_2 + a_1 + active5 + a_2
                    skill_active_list_l5.append(str5)
                if ls_skills_other_6[i] == ls_skills_other_list_f[i19]:
                    str6 = s_1 + skill3 + s_2 + a_1 + active6 + a_2
                    skill_active_list_l5.append(str6)
                if ls_skills_other_7[i] == ls_skills_other_list_f[i19]:
                    str7 = s_1 + skill3 + s_2 + a_1 + active7 + a_2
                    skill_active_list_l5.append(str7)
                if ls_skills_other_8[i] == ls_skills_other_list_f[i19]:
                    str8 = s_1 + skill4 + s_2 + a_1 + active8 + a_2
                    skill_active_list_l5.append(str8)
                if ls_skills_other_9[i] == ls_skills_other_list_f[i19]:
                    str9 = s_1 + skill4 + s_2 + a_1 + active9 + a_2
                    skill_active_list_l5.append(str9)
                if ls_skills_other_10[i] == ls_skills_other_list_f[i19]:
                    str10 = s_1 + skill5 + s_2 + a_1 + active10 + a_2
                    skill_active_list_l5.append(str10)
                if ls_skills_other_11[i] == ls_skills_other_list_f[i19]:
                    str11 = s_1 + skill5 + s_2 + a_1 + active11 + a_2
                    skill_active_list_l5.append(str11)
                if ls_skills_other_12[i] == ls_skills_other_list_f[i19]:
                    str12 = s_1 + skill6 + s_2 + a_1 + active12 + a_2
                    skill_active_list_l5.append(str12)
                if ls_skills_other_13[i] == ls_skills_other_list_f[i19]:
                    str13 = s_1 + skill6 + s_2 + a_1 + active13 + a_2
                    skill_active_list_l5.append(str13)
                if ls_skills_other_14[i] == ls_skills_other_list_f[i19]:
                    str14 = s_1 + skill7 + s_2 + a_1 + active14 + a_2
                    skill_active_list_l5.append(str14)
                if ls_skills_other_15[i] == ls_skills_other_list_f[i19]:
                    str15 = s_1 + skill7 + s_2 + a_1 + active15 + a_2
                    skill_active_list_l5.append(str15)
                if ls_skills_other_16[i] == ls_skills_other_list_f[i19]:
                    str16 = s_1 + skill8 + s_2 + a_1 + active16 + a_2
                    skill_active_list_l5.append(str16)
                if ls_skills_other_17[i] == ls_skills_other_list_f[i19]:
                    str17 = s_1 + skill8 + s_2 + a_1 + active17 + a_2
                    skill_active_list_l5.append(str17)
                if ls_skills_other_18[i] == ls_skills_other_list_f[i19]:
                    str18 = s_1 + skill9 + s_2 + a_1 + active18 + a_2
                    skill_active_list_l5.append(str18)
                if ls_skills_other_19[i] == ls_skills_other_list_f[i19]:
                    str19 = s_1 + skill9 + s_2 + a_1 + active19 + a_2
                    skill_active_list_l5.append(str19)

                if len(skill_active_list) == 5:
                    r = e + 7
                    o = y + 7
                elif len(skill_active_list) == 4:
                    r = e + 6
                    o = y + 6
                elif len(skill_active_list) == 3:
                    r = e + 5
                    o = y + 5
                elif len(skill_active_list) == 2:
                    r = e + 4
                    o = y + 4
                elif len(skill_active_list) == 1:
                    r = e + 3
                    o = y + 3
                else:
                    r = e + 2
                    o = y + 2

                for item_1 in skill_active_list_l5:  # 5个
                    if len(item_1) != 0:
                        skill_active_list_2.append(item_1)
                skill_active_list_2 = "\n".join(skill_active_list_2)  # 换行符连接
                print(len(skill_active_list_2))
                print('低分列表长度：' + str(len(skill_active_list_l5)))
                # r_1 = q + str(r) + w + str(o)
                r_1 = q + str(r - 1) + w + str(r + len(skill_active_list_l5) - 2)
                print(r_1)
                sheet.merge_range(r_1, skill_active_list_2, cell_format_content_skill)
            # 不显示
            else:
                if len(skill_active_list) == 5:
                    r = e + 7
                    o = y + 7
                elif len(skill_active_list) == 4:
                    r = e + 6
                    o = y + 6
                elif len(skill_active_list) == 3:
                    r = e + 5
                    o = y + 5
                elif len(skill_active_list) == 2:
                    r = e + 4
                    o = y + 4
                elif len(skill_active_list) == 1:
                    r = e + 3
                    o = y + 3
                else:
                    r = e + 2
                    o = y + 2

                r_1 = q + str(r) + w + str(r)
                print(r_1)
                sheet.merge_range(r_1, '由于你的行为得分同分数较多，此处不做显示', cell_format_content_skill)

            # 【请注意】
            if len(skill_active_list_l5) == 5:
                n = d + 7
            elif len(skill_active_list_l5) == 4:
                n = d + 6
            elif len(skill_active_list_l5) == 3:
                n = d + 5
            elif len(skill_active_list_l5) == 2:
                n = d + 4
            elif len(skill_active_list_l5) == 1:
                n = d + 3
            else:
                n = d + 2

            loc_other_skill_1 = q + str(n + 1) + w1 + str(n + 1)
            print(loc_other_skill_1)
            sheet.merge_range(loc_other_skill_1, "【请注意】360度评估的一个重要价值是你可以通过比较自己和他人评分,发现自己被忽视的盲点和未开发的潜能。",
                              cell_format_title1)
            loc_other_skill_1_1 = q + str(n + 2) + w1 + str(n + 2)
            sheet.merge_range(loc_other_skill_1_1, "这对你未来的发展至关重要。",
                              cell_format_title1)
            loc_other_skill_2 = q + str(n + 3) + w1 + str(n + 3)
            sheet.merge_range(loc_other_skill_2, "盲点：意味着你大大高估了自己在某个方面的能力，在该能力上存在过度自信的风险，从而影响你的在工作的表现。",
                              cell_format_content_skill)
            loc_other_skill_3 = q + str(n + 4) + w1 + str(n + 4)
            sheet.merge_range(loc_other_skill_3, "潜能：意味着你大大低估了自己在某个方面的能力，对自己的优势不够了解或缺乏自信，这将同样影响你的工作表现。",
                              cell_format_content_skill)
            loc_other_skill_4 = q + str(n + 6) + w + str(n + 6)
            sheet.merge_range(loc_other_skill_4, "识别方法", cell_format_title1)
            loc_other_skill_5 = q + str(n + 7) + w1 + str(n + 7)
            sheet.merge_range(loc_other_skill_5,
                              "请对照上表（领导技能得分明细表），1）找到“他评均分”列和“自评分数”列，2）对比每两列的颜色，若“他评均分”",
                              cell_format_content_skill)
            loc_other_skill_6 = q + str(n + 8) + w1 + str(n + 8)
            sheet.merge_range(loc_other_skill_6,
                              "为浅色，“自评分数”为深色，说明在该能力上你存在盲点，需要加强自我认知，做出适当的调整；若”他评均分”为深",
                              cell_format_content_skill)
            loc_other_skill_7 = q + str(n + 9) + w1 + str(n + 9)
            sheet.merge_range(loc_other_skill_7,
                              "色，“自评分数”为浅色，说明在该能力上你存在未开发的潜能，在增加自信的同时可以思考如何更好发挥其作用。",
                              cell_format_content_skill)
            # loc_other_skill_8 = q + str(n + 9) + w1 + str(n + 9)
            # sheet.merge_range(loc_other_skill_8,
            #                   "",
            #                   cell_format_content_skill)
            # loc_other_skill_6 = q + str(n + 9) + w1 + str(n + 10)
            # sheet.merge_range(loc_other_skill_6,
            #                   "",
            #                   cell_format_content_skill)
            sheet.write(q + str(n + 11), '3、开放式反馈', cell_format_title)
            sheet.write(q + str(n + 12), '上级', cell_format_title)
            sheet.write(q + str(n + 13), '1. 上级认为你的优势有哪些？', cell_format_content)

            print(information_mis[i])
            index_row = sheet_answer[mis_id == information_mis[i]].index.tolist()  # 该mis所在的行号列表
            print(index_row)
            role_superior = []
            role_superior_dis = []
            # 上级
            # 上级认为你的优势
            n_sa_l = 0
            for inde_r in index_row:
                if role[inde_r] == '上级':
                    role_superior.append(advantage[inde_r])
                    print(len(advantage[inde_r]))
                    if len(advantage[inde_r]) > 120:
                        print(advantage[inde_r])
                        n_sa_l = n_sa_l + 6
                    elif len(advantage[inde_r]) > 100:
                        print(advantage[inde_r])
                        n_sa_l = n_sa_l + 5
                    elif len(advantage[inde_r]) > 80:
                        print(advantage[inde_r])
                        n_sa_l = n_sa_l + 4
                    elif len(advantage[inde_r]) > 60:
                        print(advantage[inde_r])
                        n_sa_l = n_sa_l + 3
                    elif len(advantage[inde_r]) <= 60:
                        print(advantage[inde_r])
                        n_sa_l = n_sa_l + 1

            role_superior_len = len(role_superior)
            print(role_superior_len)
            print(n_sa_l)
            role_superior_1 = "\n".join(role_superior)  # 换行符连接
            loc_role_superior = q + str(n + 14) + w + str(n + 14 + (role_superior_len - 1) + n_sa_l)
            sheet.merge_range(loc_role_superior, role_superior_1, cell_format_content_skill)
            # 上级认为你的不足
            sheet.write(q + str(n + 14 + (role_superior_len - 1) + n_sa_l + 1), '2. 上级认为你的不足有哪些？对你的发展建议是什么?',
                        cell_format_content)
            print(q + str(n + 14 + (role_superior_len - 1) + n_sa_l + 1))
            n_sd_l = 0
            for inde_r in index_row:
                if role[inde_r] == '上级':
                    role_superior_dis.append(disadvantage[inde_r])
                    print(len(disadvantage[inde_r]))
                    if len(disadvantage[inde_r]) > 200:
                        print(disadvantage[inde_r])
                        n_sd_l = n_sd_l + 5
                    elif len(disadvantage[inde_r]) > 150:
                        print(disadvantage[inde_r])
                        n_sd_l = n_sd_l + 4
                    elif len(disadvantage[inde_r]) > 100:
                        print(disadvantage[inde_r])
                        n_sd_l = n_sd_l + 3
                    elif len(disadvantage[inde_r]) > 60:
                        print(disadvantage[inde_r])
                        n_sd_l = n_sd_l + 2
                    elif len(disadvantage[inde_r]) <= 60:
                        print(disadvantage[inde_r])
                        n_sd_l = n_sd_l + 1
            print(role_superior_dis)
            role_superior_dis_len = len(role_superior_dis)
            print(role_superior_dis_len)
            print(n_sd_l)
            role_superior_2 = "\n".join(role_superior_dis)  # 换行符连接 (全为数字时出现报错)
            loc_role_superior_dis = q + str(n + 14 + (role_superior_len - 1) + n_sa_l + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (role_superior_dis_len - 1) + n_sd_l + 1)
            print(loc_role_superior_dis)
            sheet.merge_range(loc_role_superior_dis, role_superior_2, cell_format_content_skill)

            # 同级伙伴
            role_companion = []
            role_companion_dis = []
            sheet.write(q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (role_superior_dis_len - 1) + n_sd_l + 1 + 2),
                        '同级/合作伙伴 ',
                        cell_format_title)
            sheet.write(q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1),
                        '1. 同级/合作伙伴认为你的优势有哪些？',
                        cell_format_content)
            # 同级认为你的优势
            n_ca_l = 0
            for inde_r in index_row:
                if role[inde_r] == '同级/合作伙伴':
                    role_companion.append(advantage[inde_r])
                    print(len(advantage[inde_r]))
                    if len(advantage[inde_r]) > 200:
                        print(advantage[inde_r])
                        n_ca_l = n_ca_l + 5
                    elif len(advantage[inde_r]) > 150:
                        print(advantage[inde_r])
                        n_ca_l = n_ca_l + 4
                    elif len(advantage[inde_r]) > 100:
                        print(advantage[inde_r])
                        n_ca_l = n_ca_l + 3
                    elif len(advantage[inde_r]) > 60:
                        print(advantage[inde_r])
                        n_ca_l = n_ca_l + 2
                    elif len(advantage[inde_r]) <= 60:
                        print(advantage[inde_r])
                        n_ca_l = n_ca_l
            print(n_ca_l)
            role_companion_len = len(role_companion)
            if role_companion_len <= n_ca_l:
                n_ca_l = n_ca_l
            else:
                n_ca_l = role_companion_len
            print('几位同级：' + str(role_companion_len))
            role_companion_1 = "\n".join(role_companion)  # 换行符连接 (全为数字时出现报错)
            loc_role_companion = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l)
            sheet.merge_range(loc_role_companion, role_companion_1, cell_format_content_skill)
            # 同级认为你的不足
            sheet.write(q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1),
                        '2. 同级/合作伙伴认为你的不足有哪些？对你的发展建议是什么? ', cell_format_content)
            n_cd_l = 0
            for inde_r in index_row:
                if role[inde_r] == '同级/合作伙伴':
                    role_companion_dis.append(disadvantage[inde_r])
                    print(len(disadvantage[inde_r]))
                    if len(disadvantage[inde_r]) > 200:
                        print(disadvantage[inde_r])
                        n_cd_l = n_cd_l + 5
                    elif len(disadvantage[inde_r]) > 150:
                        print(disadvantage[inde_r])
                        n_cd_l = n_cd_l + 4
                    elif len(disadvantage[inde_r]) > 100:
                        print(disadvantage[inde_r])
                        n_cd_l = n_cd_l + 3
                    elif len(disadvantage[inde_r]) > 60:
                        print(disadvantage[inde_r])
                        n_cd_l = n_cd_l + 2
                    elif len(disadvantage[inde_r]) <= 60:
                        print(disadvantage[inde_r])
                        n_cd_l = n_cd_l
            role_companion_dis_len = len(role_companion_dis)
            if role_companion_dis_len <= n_cd_l:
                n_cd_l = n_cd_l
            else:
                n_cd_l = role_companion_dis_len
            role_companion_2 = "\n".join(role_companion_dis)  # 换行符连接 (全为数字时出现报错)
            loc_role_companion_dis = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1)
            sheet.merge_range(loc_role_companion_dis, role_companion_2, cell_format_content_skill)

            # 下级
            role_lower = []
            role_lower_dis = []
            sheet.write(q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2),
                        '下级 ', cell_format_title)
            sheet.write(q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1),
                        '1. 下级认为你的优势有哪些？', cell_format_content)
            # 下级认为你的优势
            n_la_l = 0
            for inde_r in index_row:
                if role[inde_r] == '下级':
                    role_lower.append(advantage[inde_r])
                    if len(advantage[inde_r]) > 200:
                        print(advantage[inde_r])
                        n_la_l = n_la_l + 5
                    elif len(advantage[inde_r]) > 150:
                        print(advantage[inde_r])
                        n_la_l = n_la_l + 4
                    elif len(advantage[inde_r]) > 100:
                        print(advantage[inde_r])
                        n_la_l = n_la_l + 3
                    elif len(advantage[inde_r]) > 60:
                        print(advantage[inde_r])
                        n_la_l = n_la_l + 2
                    elif len(advantage[inde_r]) <= 60:
                        print(advantage[inde_r])
                        n_la_l = n_la_l
            role_lower_len = len(role_lower)
            if role_lower_len <= n_la_l:
                n_la_l = n_la_l
            else:
                n_la_l = role_lower_len
            role_lower_1 = "\n".join(role_lower)  # 换行符连接 (全为数字时出现报错)
            loc_role_lower = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 2)
            sheet.merge_range(loc_role_lower, role_lower_1, cell_format_content_skill)

            # 下级认为你的不足
            sheet.write(q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3),
                        '2. 下级认为你的不足有哪些？对你的发展建议是什么? ', cell_format_content)
            n_ld_l = 0
            for inde_r in index_row:
                if role[inde_r] == '下级':
                    role_lower_dis.append(disadvantage[inde_r])
                    if len(disadvantage[inde_r]) > 200:
                        print(disadvantage[inde_r])
                        n_ld_l = n_ld_l + 5
                    elif len(disadvantage[inde_r]) > 150:
                        print(disadvantage[inde_r])
                        n_ld_l = n_ld_l + 4
                    elif len(disadvantage[inde_r]) > 100:
                        print(disadvantage[inde_r])
                        n_ld_l = n_ld_l + 3
                    elif len(disadvantage[inde_r]) > 60:
                        print(disadvantage[inde_r])
                        n_ld_l = n_ld_l + 2
                    elif len(disadvantage[inde_r]) <= 60:
                        print(disadvantage[inde_r])
                        n_ld_l = n_ld_l
            role_lower_dis_len = len(role_lower_dis)
            if role_lower_dis_len <= n_ld_l:
                n_ld_l = n_ld_l
            else:
                n_ld_l = role_lower_dis_len
            role_lower_2 = "\n".join(role_lower_dis)  # 换行符连接 (全为数字时出现报错)
            loc_role_dis_lower = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2)
            sheet.merge_range(loc_role_dis_lower, role_lower_2, cell_format_content_skill)

            # 三、发展建议
            sheet.write('A' + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2),
                        '三、发展建议', cell_format_title)
            loc_3 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1)
            sheet.merge_range(loc_3, '基于领导技能和工作理念的评估结果，结合开放式反馈中的内容，建议你与上级、HR一起探讨接下来的改进方向和具体行',
                              cell_format_content_skill)
            #
            loc_3_1 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1)
            sheet.merge_range(loc_3_1, '动计划：',
                              cell_format_content_skill)
            loc_3_2 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_3_2, '1.聚焦你希望重点发展的领域。', cell_format_title)
            str_suggesstion1 = '本报告从不同侧面显示了你可以提升的地方, 但并不是每一个都同等重要。你需要结合个人目标和上级期望, 确定发展策略， '
            loc_4 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_4, str_suggesstion1, cell_format_content_skill)
            #
            str_suggesstion1_1 = '最终确定2-3个对改进绩效和个人发展最重要的方面，作为接下来半年或一年的发展目标。'
            loc_4_1 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_4_1, str_suggesstion1_1, cell_format_content_skill)
            #
            loc_4_1_0 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_4_1_0, '2.思考合适的发展方式。', cell_format_title)
            # # loc_4_2 = q + str(
            # #     n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2 + 8) + w + str(
            # #     n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2 + 8)
            # # sheet.merge_range(loc_4_2, str_suggesstion1_2, cell_format_content_skill)
            #
            str_suggesstion2 = '确定目标后，你需要思考什么样的发展方式最适合自己。成人学习理论认为，最有效的发展70%来自于日常的实践 （如挑'
            loc_5 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_5, str_suggesstion2, cell_format_content_skill)
            #
            str_suggesstion2_1 = '战性任务、跨团队项目 ）自于日常的实践（如挑战性任务、跨团队项目），20%来自身边的人 （如上级的辅导），10%'
            loc_5_1 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_5_1, str_suggesstion2_1, cell_format_content_skill)
            #
            str_suggesstion2_2 = '来自课堂（如培训)。 确保你的计划包括不同的发展方式。'
            loc_5_2 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_5_2, str_suggesstion2_2, cell_format_content_skill)
            #
            str_suggesstion3 = '3.撰写你的个人发展计划初稿。'
            loc_6 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_6, str_suggesstion3, cell_format_title)
            #
            str_suggesstion3_1 = '使用附录中的表格开始撰写你的个人发展计划。首先思考你的改变可能带来的影响,然后描述你计划采取的行动、需要的'
            loc_6_1 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_6_1, str_suggesstion3_1, cell_format_content_skill)
            #
            str_suggesstion3_2 = '资源和时间表，最后还需要考虑可能遇到的障碍及如何跟进计划。'
            loc_6_2 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_6_2, str_suggesstion3_2, cell_format_content_skill)
            #
            str_suggesstion4 = '4.准备好和上级、HR的讨论。'
            loc_7 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_7, str_suggesstion4, cell_format_title)
            #
            str_suggesstion4_1 = '主动和上级、HR讨论你的个人发展计划，听取上级、HR的额外反馈和建议， 并基于新的输入更新发展计划。'
            loc_7_1 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_7_1, str_suggesstion4_1, cell_format_content_skill)
            #
            str_suggesstion5 = '5.按照计划跟进。'
            loc_8 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1)
            sheet.merge_range(loc_8, str_suggesstion5, cell_format_title)
            #
            str_suggesstion6 = '根据计划中约定的跟进机制，定期和上级、HR讨论进展，获得更多的反馈和指导。'
            loc_9 = q + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 21) + w + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 21)
            sheet.merge_range(loc_9, str_suggesstion6, cell_format_content_skill)
            #
            sheet.write('A' + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 22),
                        '四、附录', cell_format_title)
            sheet.write('B' + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 23),
                        '个人发展计划表格', cell_format_title)
            #
            # 附录
            sheet.insert_image('B' + str(
                n + 14 + role_superior_len - 1 + + n_sa_l + 1 + 1 + (
                            role_superior_dis_len - 1) + n_sd_l + 1 + 2 + 1 + 1 + n_ca_l + 1 + 1 + n_cd_l + 1 + 2 + 1 + 1 + n_la_l + 3 + 1 + n_ld_l + 25),
                               fulu + '附录1.png', {'x_scale': 0.8, 'y_scale': 0.8})  # , {'x_scale': 0.8, 'y_scale': 0.8}
            # sheet.write('B' + str(n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2+22),'待发展方向',property_table1)
            # sheet.write('C' + str(n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2+22),'期待结果',property_table1)
            # sheet.write('D' + str(n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2+22),'行动计划',property_table1)
            # sheet.write('E' + str(n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2+22),'潜在障碍',property_table1)
            # sheet.write('F' + str(n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2+22),'需要的资源/支持',property_table1)
            # sheet.write('G' + str(n + 14 + role_superior_len + 1 + role_superior_dis_len + 3 + role_companion_len + 2 + role_companion_dis_len + 3 + role_lower_len + 2+22),'开始/完成日期',property_table1)
            #

            # 将全部的Excel文件存储（根据mis号进行文件命名，存储到同一文件夹下，image等分类单独存储到文件夹中）

            book.close()
