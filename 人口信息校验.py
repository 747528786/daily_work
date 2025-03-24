import time
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from numpy import nan
from identiti_card import IdentityCard
from collections import Counter


class SheetNameError(Exception):
    pass


def get_path(file_name: str) -> str:
    # 选择文件路径
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title=file_name)
    return path


def data_cleaning(path):
    miss_sheet_list = []
    print('开始读取数据')
    df = pd.ExcelFile(path)
    print(df.sheet_names)
    if '明细' not in df.sheet_names:
        miss_sheet_list.append('没有叫明细的工作表')
    if '房屋信息' not in df.sheet_names:
        miss_sheet_list.append('没有叫房屋信息的工作表')
    if '人房信息' not in df.sheet_names:
        miss_sheet_list.append('没有叫人房信息的工作表')
    if miss_sheet_list:
        raise SheetNameError(miss_sheet_list)
    df_file = df.parse('明细')
    df_house = df.parse('房屋信息')
    df_db = df.parse('人房信息')
    df_file['证件号码'] = df_file['证件号码'].str.upper()
    df_db['证件号码'] = df_db['证件号码'].str.upper()

    '''
    计算重复数据条数并删除
    '''
    duplicate_data = df_file.duplicated().sum()
    print(f'有{duplicate_data}条重复数据，已进行删除')
    df_file.drop_duplicates(inplace=True)
    df_file['备注'] = ''

    '''
    定义必填项，校验必填项是否为空
    '''
    print('校验必填项是否为空')
    not_na_list = ["街道", "社区", "小区", "楼幢", "房间号", "姓名",
                   "证件类型", "证件号码", "出生日期", "性别", "国家及地区",
                   "人员类型", "人员状态", "居住状态"]
    for col in not_na_list:
        df_file.loc[df_file[col].isna(), ['备注']] += '丨' + col + '列不能为空'

    '''
    定义数据字典，校验输入规则
    '''
    print('校验是否符合输入规则')
    id_type_list = ["居民身份证", "香港特区护照/身份证明", "澳门特区护照/身份证明", "台湾居民来往大陆通行证", "永久居留证", "归国华侨证", "外国人护照", "其他"]
    politic_status_list = ["中共党员", "共青团员", "民主党派", "群众", "未知", nan]
    edu_list = ["博士", "硕士", "本科", "大专", "高中", "初中", "小学及以下", "未知", "中专", nan]
    people_type_list = ["户籍", "流动", "非户籍常住"]
    people_status_list = ["正常", "死亡", "未知"]
    live_status_list = ["租住", "自住"]
    marry_list = ["已婚", "未婚", "离异", "未知", nan]
    sex_list = ["男", "女"]
    dict1 = {"证件类型": "id_type",
             "政治面貌": "politic_status",
             "文化程度": "edu",
             "人员类型": "people_type",
             "人员状态": "people_status",
             "居住状态": "live_status",
             "婚姻状况": "marry",
             "性别": "sex"}
    for key, value in dict1.items():
        # print(eval(value + '_list'))
        df_file.loc[~df_file[key].isin(eval(value + '_list')), ['备注']] += '丨' + key + '不符合输入规则'

    '''
    填充楼幢房间号中的幢、室
    方法一：先把幢和室都替换为空，再统一加上
    df_file['楼幢'].replace('幢', '')
    df_file['楼幢'] = df_file['楼幢'].astype(str) + '幢'
    df_file['房间号'].replace('室', '')
    df_file['房间号'] = df_file['房间号'].astype(str) + '室'
    '''
    print('楼幢、房间号填充幢、室')
    df_file['楼幢'] = df_file['楼幢'].apply(lambda x: str(x) + '幢' if '幢' not in str(x) else x)
    df_file['房间号'] = df_file['房间号'].apply(lambda x: str(x) + '室' if '室' not in str(x) else x)

    '''
    校验身份证号码
    '''
    print('校验身份证号码')
    df_file['身份证号码校验'] = df_file['证件号码'].map(lambda x: IdentityCard.is_id_card(x))

    '''
    校验联系方式列，不能重复10次及以上
    Counter计算联系方式列各个值出现次数，返回一个字典
    筛选字典值>=10次的，即为手机号重复10次及以上
    '''
    print('校验联系方式')
    # df_file['len_concat'] = df_file['联系方式'].map(lambda x: len(str(x)) if x is not None else x)
    result = Counter(df_file['联系方式'])
    result_filter = {key: val for key, val in result.items() if val >= 10}
    # print(result_filter)

    if result_filter:
        list_phone = []
        for key, value in result_filter.items():
            temp = key
            list_phone.append(temp)
        # print(list_phone)
        df_file.loc[df_file['联系方式'].isin(list_phone), ['备注']] += '丨' + '联系方式重复10次及以上'

    '''
    校验房屋信息，备注列输出系统中无此房间号
    校验同一房号下居住人数是否大于10人
    '''
    print('校验系统中是否有此房间号')
    df_file.loc[~df_file['小区'].isin(df_house['小区']), ['备注']] += '丨系统中无此小区'
    df_file['小区楼幢房间号'] = df_file['小区'].map(str) + df_file['楼幢'].map(str) + df_file['房间号'].map(str)
    df_house['小区楼幢房间号'] = df_house['小区'].map(str) + df_house['楼幢'].map(str) + df_house['房间号'].map(str)
    df_file.loc[~df_file['小区楼幢房间号'].isin(df_house['小区楼幢房间号']), ['备注']] += '丨系统中无此房间号'
    print('校验同一房号下居住人数是否大于10人')
    house_count = Counter(df_file['小区楼幢房间号'])
    house_count_filter = {key: val for key, val in house_count.items() if val > 10}
    if house_count_filter:
        list_house = []
        for key, value in house_count_filter.items():
            temp = key
            list_house.append(temp)
        df_file.loc[df_file['小区楼幢房间号'].isin(list_house), '备注'] += '丨同一房号下居住人数大于10人'
    df_file.drop(['小区楼幢房间号'], axis=1, inplace=True)

    '''
    校验同一房间下姓名相同、证件号码重复
    '''
    print('校验同一房间下姓名相同、证件号码重复')
    list_count = ['姓名', '证件号码']
    for i in list_count:
        df_file['小区楼幢房间号' + i] = \
            df_file['小区'].map(str) + df_file['楼幢'].map(str) + df_file['房间号'].map(str) + df_file[i].map(str)
        count_name = Counter(df_file['小区楼幢房间号' + i])
        count_name_filter = {key: val for key, val in count_name.items() if val > 1}
        if count_name_filter:
            list_name = []
            for key, value in count_name_filter.items():
                temp = key
                list_name.append(temp)
            df_file.loc[df_file['小区楼幢房间号' + i].isin(list_name), ['备注']] += '丨同一房号下' + i + '重复'
            df_file.drop(['小区楼幢房间号' + i], axis=1, inplace=True)
        else:
            df_file.drop(['小区楼幢房间号' + i], axis=1, inplace=True)

    '''
    校验是否为重复数据
    '''
    print('校验是否重复数据')
    df_file['小区楼幢房间号证件号码'] = \
        df_file['小区'].map(str) + df_file['楼幢'].map(str) + df_file['房间号'].map(str) + df_file['证件号码'].map(str)
    df_db['小区楼幢房间号证件号码'] = \
        df_db['小区'].map(str) + df_db['楼幢'].map(str) + df_db['房间号'].map(str) + df_db['证件号码'].map(str)
    df_file.loc[df_file['小区楼幢房间号证件号码'].isin(df_db['小区楼幢房间号证件号码']), ['备注']] += '丨重复数据'
    df_file.drop(['小区楼幢房间号证件号码'], axis=1, inplace=True)
    return df_file


if __name__ == '__main__':
    file_path = get_path('选择需校验的文件')
    new_path = file_path.replace(os.path.basename(os.path.normpath(file_path)), "校验后.xlsx")
    data_cleaning(path=file_path).to_excel(new_path, index=False)
    print('校验完成')
    time.sleep(20)
