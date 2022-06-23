# coding = utf-8
import csv
import json
import logging
import pandas as pd
from loguru import logger
from conf.file_conf import data_path  # 路径查找

#  传入文件
org_file_name = data_path + '/org_20211231_1.txt'  # 组织架构文件
dx_f1 = '/DX_Q12_dd_label.txt'
dx_f2 = '/DX_Q34_dd_label.txt'
dx_file_name = data_path  # 组织架构文件
jf_file_name = data_path + 'job_family_20211231_2.txt'
ona_label = data_path + 'dd_ona_label.txt'


# 部门架构
def org_name_dept(org_file_name):
    df_org = pd.read_csv(org_file_name, encoding='utf-8',
                         dtype={'dept_code': str, 'mapping_top_04_dept_code': str,
                                'mapping_top_03_dept_code': str,
                                'mapping_top_05_dept_code': str,
                                'mapping_top_06_dept_code': str,
                                'mapping_top_07_dept_code': str,
                                'dept_code_path': str}, sep='\t')
    # df_org = df_org[df_org["dept_code_path"].str.contains("104493")]  # 到店事业群
    return df_org


df_org = org_name_dept(org_file_name)


# 数据量较大，用于合并数据
def dx_messge_merge(dx_file_name):
    DX_Q12_dd = pd.read_csv(dx_file_name + dx_f1, sep='\t',
                            dtype={'emp_code': str, 'to_emp_code': str, 'dept_code': str,
                                   'to_dept_code': str})  # csv文件的分句结果好于txt
    DX_Q34_dd = pd.read_csv(dx_file_name + dx_f2, sep='\t',
                            dtype={'emp_code': str, 'to_emp_code': str, 'dept_code': str,
                                   'to_dept_code': str})  # csv文件的分句结果好于txt
    DX_Q1234_dd = pd.concat([DX_Q12_dd, DX_Q34_dd])  # 全年四个季度数据合并

    return DX_Q1234_dd


DX_Q1234_dd = dx_messge_merge(dx_file_name)


# 分ba和ab进行识别，防止去掉结点
def DX_BA_sr(DX_Q1234_dd):
    # 去除无效数据
    # 删除符合条件的指定行，并替换原始df
    # DX_Q1234_dd.drop(DX_Q1234_dd[(DX_Q1234_dd.qua_r_num == 0)].index,inplace = True)
    # DX_1234_new = DX_Q1234_dd.query('r_day_cnt != 0')  # 筛选r
    DX_Q1234_dd['key_value'] = (DX_Q1234_dd['s_day_cnt'] + DX_Q1234_dd['r_day_cnt']) / 2  # 新增Key列
    DX_Q1234_dd['A-to-B'] = DX_Q1234_dd['emp_code'] + '' + DX_Q1234_dd['to_emp_code']  # A to B
    DX_Q1234_dd['B-to-A'] = DX_Q1234_dd['to_emp_code'] + '' + DX_Q1234_dd['emp_code']  # B to A
    DX_Q1234_dd['key_value'] = DX_Q1234_dd['key_value'].map(lambda x: str(x))
    DX_Q1234_dd['A-to-B_key'] = DX_Q1234_dd["A-to-B"].str.cat(DX_Q1234_dd["key_value"], sep="")
    DX_Q1234_dd['B-to-A_key'] = DX_Q1234_dd["B-to-A"].str.cat(DX_Q1234_dd["key_value"], sep="")
    # DX_1234_new = DX_Q1234_dd.query('A-to-B != "0"')
    # A_B_index_label = pd.DataFrame(DX_Q1234_dd,columns=['A-to-B','key_value'])  # 创建新的数据框
    # cha
    DX_1234_1 = DX_Q1234_dd[DX_Q1234_dd['A-to-B_key'].isin(DX_Q1234_dd['B-to-A_key'])]  # AB、BA
    DX_1234_2 = DX_Q1234_dd[~DX_Q1234_dd['A-to-B_key'].isin(DX_Q1234_dd['B-to-A_key'])]  # 非AB、BA
    # 列互换
    DX_1234_2[['emp_code', 'to_emp_code']] = DX_1234_2[['to_emp_code', 'emp_code']]
    # DX_1234_2[['emp_mis_name', 'to_emp_mis_name']] = DX_1234_2[['to_emp_mis_name', 'emp_mis_name']]
    DX_1234_2[['dept_code', 'to_dept_code']] = DX_1234_2[['to_dept_code', 'dept_code']]
    DX_1234_2[['s_day_cnt', 'r_day_cnt']] = DX_1234_2[['r_day_cnt', 's_day_cnt']]
    # DX_1234_2[['s_week_cnt', 'r_week_cnt']] = DX_1234_2[['r_week_cnt', 's_week_cnt']]
    # DX_1234_2[['s_month_cnt', 'r_month_cnt']] = DX_1234_2[['r_month_cnt', 's_month_cnt']]
    # DX_1234_2[['qua_s_num', 'qua_r_num']] = DX_1234_2[['qua_r_num', 'qua_s_num']]
    DX_1234_2 = DX_1234_2[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 's_day_cnt', 'r_day_cnt']]
    DX_1234_new = pd.concat([DX_Q1234_dd, DX_1234_2])  # 上述两种情况合并
    #
    return DX_1234_new


DX_1234_new = DX_BA_sr(DX_Q1234_dd)


# 关联序列(时间Label)
def jf(jf_file_name, DX_1234_new):
    jf = pd.read_csv(jf_file_name, sep='\t', dtype={'emp_code': str})
    DX_jf = pd.merge(
        DX_1234_new,
        jf,
        left_on=['emp_code', 'Label'],
        right_on=['emp_code', 'time_label'],
        how='left'
    )
    DX_jf = DX_jf.rename(columns={'emp_code_x': 'emp_code'})  # 修改列名
    DX_jf = DX_jf[
        ['emp_code', 'to_emp_code', 'emp_mis_name', 'to_emp_mis_name', 'dept_code', 'to_dept_code', 's_day_cnt',
         'r_day_cnt', 'Label', 'emp_job_family_desc']]
    # 's_week_cnt', 'r_week_cnt','s_month_cnt', 'r_month_cnt', 'qua_s_num', 'qua_r_num',
    DX_jf = DX_jf.query('emp_job_family_desc == "BA"')
    DX_jf.drop_duplicates(subset=None, keep='first', inplace=True)
    return DX_jf


DX_jf = jf(jf_file_name, DX_1234_new)


# 关联各级部门
def BA_dept(DX_jf, df_org):
    # 筛选出发出者为“到店事业群”
    BA_dept1 = pd.merge(
        DX_jf,
        df_org,
        left_on='dept_code',
        right_on='dept_code',
        how='left'
    )
    BA_dept1 = BA_dept1.rename(columns={'dept_code_x': 'dept_code'})
    # BA_dept1 = BA_dept1.dropna(subset=["dept_code_path"])  # 删除为NAN的行
    BA_dept1 = BA_dept1[BA_dept1["dept_code_path"].str.contains("104493")]  # 到店事业群
    BA_dept1 = BA_dept1[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 's_day_cnt', 'r_day_cnt']]

    # 筛选出接收者为“到店事业群”
    BA_dept2 = pd.merge(
        BA_dept1,
        df_org,
        left_on='to_dept_code',
        right_on='dept_code',
        how='left'
    )
    BA_dept2 = BA_dept2.rename(columns={'dept_code_x': 'dept_code'})
    # BA_dept2 = BA_dept2.dropna(subset=["dept_code_path"])  # 删除为NAN的行
    BA_dept = BA_dept2[BA_dept2["dept_code_path"].str.contains("104493")]  # 到店事业群
    BA_dept = BA_dept[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 's_day_cnt', 'r_day_cnt', 'dept_code_path',
         'dept_name_path', 'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name',
         'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code']]

    return BA_dept1,BA_dept


BA_dept1,BA_dept = BA_dept(DX_jf, df_org)
BA_dept1.to_excel('BA_dept1.xlsx', encoding='utf-8', index=False)

# 分季度拆分标签，进行特殊沟通频次的标签整理
def dx_split_q(BA_dept):
    DX_Q1_label = BA_dept.query('Label == "Q1"')
    DX_Q2_label = BA_dept.query('Label == "Q2"')
    DX_Q3_label = BA_dept.query('Label == "Q3"')
    DX_Q4_label = BA_dept.query('Label == "Q4"')

    return DX_Q1_label, DX_Q2_label, DX_Q3_label, DX_Q4_label


DX_Q1_label, DX_Q2_label, DX_Q3_label, DX_Q4_label = dx_split_q(BA_dept)

#
def dx_frequency(DX_Q1_label, DX_Q2_label, DX_Q3_label, DX_Q4_label):
    DX_Q1_label_s = DX_Q1_label.query('s_day_cnt == 1 and r_day_cnt == 1')
    DX_Q2_label_s = DX_Q2_label.query('s_day_cnt == 1 and r_day_cnt == 1')
    DX_Q3_label_s = DX_Q3_label.query('s_day_cnt == 1 and r_day_cnt == 1')
    DX_Q4_label_s = DX_Q4_label.query('s_day_cnt == 1 and r_day_cnt == 1')
    return DX_Q1_label_s, DX_Q2_label_s, DX_Q3_label_s, DX_Q4_label_s


DX_Q1_label_s, DX_Q2_label_s, DX_Q3_label_s, DX_Q4_label_s = dx_frequency(DX_Q1_label, DX_Q2_label, DX_Q3_label,
                                                                          DX_Q4_label)


# print(DX_Q1_label_s)

def dx_join(DX_Q3_label_s, DX_Q4_label_s):
    # Q1&2季度特殊沟通频次探索
    DX_Q34_dd = pd.merge(
        DX_Q3_label_s,
        DX_Q4_label_s,
        left_on=['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code'],
        right_on=['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code'],
        how='inner'
    )
    DX_Q34_dd = DX_Q34_dd.rename(columns={'emp_code_x': 'emp_code', 'to_emp_code_x': 'to_emp_code', 'Label_x': 'Label',
                                          'dept_code_x': 'dept_code', 'to_dept_code_x': 'to_dept_code',
                                          'dept_code_path_x': 'dept_code_path',
                                          'dept_name_path_x': 'dept_name_path',
                                          'mapping_top_04_dept_name_x': 'mapping_top_04_dept_name',
                                          'mapping_top_05_dept_name_x': 'mapping_top_05_dept_name',
                                          'mapping_top_06_dept_name_x': 'mapping_top_06_dept_name',
                                          'mapping_top_07_dept_name_x': 'mapping_top_07_dept_name',
                                          'mapping_top_04_dept_code_x': 'mapping_top_04_dept_code',
                                          'mapping_top_05_dept_code_x': 'mapping_top_05_dept_code',
                                          'mapping_top_06_dept_code_x': 'mapping_top_06_dept_code',
                                          'mapping_top_07_dept_code_x': 'mapping_top_07_dept_code'})
    DX_Q34_dd = DX_Q34_dd[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path', 'dept_name_path',
         'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name', 'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code']]
    return DX_Q34_dd


DX_Q34_dd = dx_join(DX_Q3_label_s, DX_Q4_label_s)


# print(DX_Q12_dd)


# 各级标签
def BA_domain_label(DX_Q34_dd, ona_label):
    ona_label = pd.read_csv(ona_label, sep='\t', dtype={'department_code ': str, 'parent_department_code ': str,
                                                        'bg_code': str})
    # 父一级标签
    BA_label = pd.merge(
        DX_Q34_dd,
        ona_label,
        left_on='mapping_top_04_dept_code',  # 修改
        right_on='parent_department_code ',
        how='left'
    )
    BA_label_1 = BA_label[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path', 'dept_name_path',
         'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name', 'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code', 'parent_realm_name']]

    BA_label_1 = BA_label_1.rename(columns={'parent_realm_name': '父一级标签'})

    # 父二级标签
    BA_label_2 = pd.merge(
        BA_label_1,
        ona_label,
        left_on='mapping_top_05_dept_code',  # 修改
        right_on='parent_department_code ',
        how='left'
    )
    BA_label_2 = BA_label_2[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path', 'dept_name_path',
         'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name', 'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code', '父一级标签', 'parent_realm_name']]
    BA_label_2 = BA_label_2.rename(columns={'parent_realm_name': '父二级标签'})

    # 父三级标签
    BA_label_3 = pd.merge(
        BA_label_2,
        ona_label,
        left_on='mapping_top_06_dept_code',  # 修改
        right_on='parent_department_code ',
        how='left'
    )

    BA_label_3 = BA_label_3[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path', 'dept_name_path',
         'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name', 'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code', '父一级标签', '父二级标签', 'parent_realm_name']]

    BA_label_3 = BA_label_3.rename(columns={'parent_realm_name': '父三级标签'})

    # 子一级标签
    BA_label_4 = pd.merge(
        BA_label_3,
        ona_label,
        left_on='mapping_top_04_dept_code',  # 修改
        right_on='department_code ',
        how='left'
    )

    BA_label_4 = BA_label_4[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path', 'dept_name_path',
         'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name', 'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code', '父一级标签', '父二级标签', '父三级标签', 'subdomain_name']]

    BA_label_4 = BA_label_4.rename(columns={'subdomain_name': '子一级标签'})

    # 子二级标签
    BA_label_5 = pd.merge(
        BA_label_4,
        ona_label,
        left_on='mapping_top_05_dept_code',  # 修改
        right_on='department_code ',
        how='left'
    )

    BA_label_5 = BA_label_5[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path', 'dept_name_path',
         'mapping_top_04_dept_name', 'mapping_top_05_dept_name', 'mapping_top_06_dept_name', 'mapping_top_07_dept_name',
         'mapping_top_04_dept_code', 'mapping_top_05_dept_code', 'mapping_top_06_dept_code',
         'mapping_top_07_dept_code', '父一级标签', '父二级标签', '父三级标签', '子一级标签', 'subdomain_name']]

    BA_label_5 = BA_label_5.rename(columns={'subdomain_name': '子二级标签'})

    # 子三级标签
    BA_label_6 = pd.merge(
        BA_label_5,
        ona_label,
        left_on='mapping_top_06_dept_code',  # 修改
        right_on='department_code ',
        how='left'
    )

    BA_label_6 = BA_label_6[
        ['emp_code', 'to_emp_code', 'dept_code', 'to_dept_code', 'Label', 'dept_code_path',
         'dept_name_path', '父一级标签', '父二级标签', '子二级标签', 'subdomain_name']]

    BA_label_6 = BA_label_6.rename(columns={'subdomain_name': '子二级标签', '子二级标签': '子一级标签'})

    # 标签清理
    BA_label_all = BA_label_6.query(
        "父一级标签 != 'NaN' or 父二级标签 != 'NaN' or 子一级标签 != 'NaN' or 子二级标签 != 'NaN'")
    # 去除重复项
    BA_label_all.drop_duplicates(subset=None, keep='first', inplace=True)

    return BA_label_all


BA_label_all = BA_domain_label(DX_Q34_dd, ona_label)
BA_count = len(BA_label_all['emp_code'].unique())
print('BA人数为：' + str(BA_count))

#  数据导出
BA_label_all.to_excel('DX_Q34_dd_11_1.xlsx', encoding='utf-8', index=False)
# print(BA_label_all)
