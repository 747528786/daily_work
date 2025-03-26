import time

import pandas as pd
from numpy import nan
import tkinter as tk
from tkinter import filedialog
from identiti_card import IdentityCard
from collections import Counter
import os
import ctypes
import sys


class SheetNameError(Exception):
    pass


class DetailError(Exception):
    pass


def get_path(file_name: str) -> str:
    # 选择文件路径
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title=file_name, filetypes=[('Excel文件', '*xlsx')])
    return path


def verify_sheet_name(df):  # 校验工作表
    """
    校验是否存在所需的三个sheet 
    :param df: DataFrame
    :return: Error
    """
    required_sheets = {'明细', '房屋信息', '人房信息'}
    miss_sheet_list = required_sheets - set(df.sheet_names)
    if miss_sheet_list:
        raise SheetNameError([f'没有找到名为"{name}"的工作表' for name in miss_sheet_list])


def col_upper(df, col='证件号码'):  # 证件号码列x转大写
    df[col] = df[col].str.upper()


def drop_duplicate_data(df):  # 计算重复数据条数并删除
    duplicate_num = df.duplicated().sum()
    if duplicate_num == 0:
        print('无重复数据')
    else:
        print(f'有{duplicate_num}条重复数据，已进行删除')
        df.drop_duplicates(inplace=True)


def verify_required_empty(df_detail: pd.DataFrame, df_room: pd.DataFrame, df_db: pd.DataFrame):  # 校验必填项是否为空
    err_list = []
    # 明细必填项
    df_detail_not_na_list = ["街道", "社区", "小区", "楼幢", "房间号", "姓名",
                             "证件类型", "证件号码", "出生日期", "性别", "国家及地区",
                             "人员类型", "人员状态", "居住状态"]
    detail_na_cols = [col for col in df_detail_not_na_list if df_detail[col].isna().any()]
    # detail_empty_col_list = df_detail.columns[df_detail[df_detail_not_na_list].isna().any()].index.tolist()
    if detail_na_cols:
        err_list.append(f'明细中{detail_na_cols}列有空值，请检查')

    # 房屋信息必填项
    room_not_na_list = ["街道", "社区", "小区", "楼幢", "房间号"]
    room_na_cols = [col for col in room_not_na_list if df_room[col].isna().any()]
    if room_na_cols:
        err_list.append(f'房屋信息中{room_na_cols}列有空值，请检查')

    # 人房信息必填项
    db_not_na_list = ["小区", "楼幢", "房间号", "姓名", "证件号码"]
    db_na_cols = [col for col in db_not_na_list if df_db[col].isna().any()]
    if db_na_cols:
        err_list.append(f'人房信息中{db_na_cols}列有空值，请检查')

    if err_list:
        raise DetailError(err_list)


def verify_required_exist(df_detail: pd.DataFrame, df_room: pd.DataFrame, df_db: pd.DataFrame):  # 校验必填项是否存在
    miss_col_list = []

    df_detail_should_set = {"街道", "社区", "小区", "楼幢", "房间号", "姓名",
                            "证件类型", "证件号码", "出生日期", "性别", "国家及地区",
                            "人员类型", "人员状态", "居住状态"}
    df_detail_header_set = set(list(df_detail))
    df_detail_miss_col_set = df_detail_should_set - (df_detail_should_set & df_detail_header_set)
    if df_detail_miss_col_set:
        miss_col_list.append('明细中缺少%s列' % (str(df_detail_miss_col_set)))

    df_room_should_set = {"街道", "社区", "小区", "楼幢", "房间号"}
    df_room_header_set = set(list(df_room))
    df_room_miss_col_set = df_room_should_set - (df_room_header_set & df_room_header_set)
    if df_room_miss_col_set:
        miss_col_list.append('房屋信息中缺少%s列' % (str(df_room_miss_col_set)))

    df_db_should_set = {"小区", "楼幢", "房间号", "姓名", "证件号码"}
    df_db_header_set = set(list(df_db))
    df_db_miss_col_set = df_db_should_set - (df_db_header_set & df_db_header_set)
    if df_db_miss_col_set:
        miss_col_list.append('人房信息中缺少%s列' % (str(df_db_miss_col_set)))
    if miss_col_list:
        raise DetailError(miss_col_list)


def verify_input_rule(df: pd.DataFrame):
    id_type_list = ["居民身份证", "香港特区护照/身份证明", "澳门特区护照/身份证明",
                    "台湾居民来往大陆通行证", "永久居留证", "归国华侨证", "外国人护照", "其他"]
    politic_status_list = ["中共党员", "共青团员", "民主党派", "群众", nan]
    edu_list = ["博士", "硕士", "本科", "大专", "高中", "初中", "小学及以下", "未知", "中专", nan]
    people_type_list = ["户籍", "流动", "非户籍常住"]
    people_status_list = ["正常", "死亡", "未知"]
    live_status_list = ["租住", "自住"]
    marry_list = ["已婚", "未婚", "离异", "未知", nan]
    sex_list = ["男", "女"]
    group_list = ["塔塔尔族", "蒙古族", "回族", "藏族", "维吾尔族", "苗族", "彝族", "壮族", "布依族", "朝鲜族", "满族", "侗族", "瑶族", "白族", "土家族",
                  "哈尼族", "哈萨克族", "傣族", "黎族", "傈僳族", "佤族", "畲族", "高山族", "拉祜族", "水族", "东乡族", "纳西族", "景颇族", "柯尔克孜族", "土族",
                  "达斡尔族", "仫佬族", "羌族", "布朗族", "撒拉族", "毛南族", "仡佬族", "锡伯族", "阿昌族", "普米族", "塔吉克族", "怒族", "乌孜别克族", "俄罗斯族",
                  "鄂温克族", "崩龙族", "保安族", "裕固族", "京族", "汉族", "独龙族", "鄂伦春族", "赫哲族", "门巴族", "珞巴族", "基诺族", "其他", "革家人族", nan]
    country_list = ["阿富汗", "阿尔巴尼亚", "阿尔及利亚", "美属萨摩亚", "安道尔", "安哥拉", "安圭拉", "南极洲", "安提瓜和巴布达", "阿根廷", "亚美尼亚", "阿鲁巴",
                    "澳大利亚", "奥地利", "阿塞拜疆", "巴哈马", "巴林", "孟加拉国", "巴巴多斯", "白俄罗斯", "比利时", "伯利兹", "贝宁", "百慕大", "不丹", "玻利维亚",
                    "波黑", "博茨瓦纳", "布维岛", "巴西", "英属印度洋领土", "文莱", "保加利亚", "布基纳法索", "布隆迪", "柬埔寨", "喀麦隆", "加拿大", "佛得角",
                    "开曼群岛", "中非", "乍得", "智利", "中国", "中国香港", "中国澳门", "中国台湾", "圣诞岛", "科科斯(基林)群岛", "哥伦比亚", "科摩罗", "刚果（布）",
                    "刚果（金）", "库克群岛", "哥斯达黎加", "科特迪瓦", "克罗地亚", "古巴", "塞浦路斯", "捷克", "丹麦", "吉布提", "多米尼克", "多米尼加共和国", "东帝汶",
                    "厄瓜多尔", "埃及", "萨尔瓦多", "赤道几内亚", "厄立特里亚", "爱沙尼亚", "埃塞俄比亚", "福克兰群岛(马尔维纳斯)", "法罗群岛", "斐济", "芬兰", "法国",
                    "法属圭亚那", "法属波利尼西亚", "法属南部领土", "加蓬", "冈比亚Gambia", "格鲁吉亚", "德国", "加纳", "直布罗陀", "希腊", "格陵兰", "格林纳达",
                    "瓜德罗普", "关岛", "危地马拉", "几内亚", "几内亚比绍", "圭亚那", "海地", "赫德岛和麦克唐纳岛", "洪都拉斯", "匈牙利", "冰岛", "印度", "印度尼西亚",
                    "伊朗", "伊拉克", "爱尔兰", "以色列", "意大利", "牙买加", "日本", "约旦", "哈萨克斯坦", "肯尼亚", "基里巴斯", "朝鲜", "韩国", "科威特",
                    "吉尔吉斯斯坦", "老挝", "拉脱维亚", "黎巴嫩", "莱索托", "利比里亚", "利比亚", "列支敦士登", "立陶宛", "卢森堡", "前南马其顿", "马达加斯加", "马拉维",
                    "马来西亚", "马尔代夫", "马里", "马耳他", "马绍尔群岛", "马提尼克", "毛里塔尼亚", "毛里求斯", "马约特", "墨西哥", "密克罗尼西亚联邦", "摩尔多瓦",
                    "摩纳哥", "蒙古", "蒙特塞拉特", "摩洛哥", "莫桑比克", "缅甸", "纳米比亚", "瑙鲁", "尼泊尔", "荷兰", "荷属安的列斯", "新喀里多尼亚", "新西兰",
                    "尼加拉瓜", "尼日尔", "尼日利亚", "纽埃", "诺福克岛", "北马里亚纳", "挪威", "阿曼", "巴基斯坦", "帕劳", "巴勒斯坦", "巴拿马", "巴布亚新几内亚",
                    "巴拉圭", "秘鲁", "菲律宾", "皮特凯恩群岛", "波兰", "葡萄牙", "波多黎各", "卡塔尔", "留尼汪", "罗马尼亚", "俄罗斯联邦", "卢旺达", "圣赫勒拿",
                    "圣基茨和尼维斯", "圣卢西亚", "圣皮埃尔和密克隆", "圣文森特和格林纳丁斯", "萨摩亚", "圣马力诺", "圣多美和普林西比", "沙特阿拉伯", "塞内加尔", "塞舌尔",
                    "塞拉利昂", "新加坡", "斯洛伐克", "斯洛文尼亚", "所罗门群岛", "索马里", "南非", "南乔治亚岛和南桑德韦奇岛", "西班牙", "斯里兰卡", "苏丹", "苏里南",
                    "斯瓦尔巴群岛", "斯威士兰", "瑞典", "瑞士", "叙利亚", "塔吉克斯坦", "坦桑尼亚", "泰国", "多哥", "托克劳", "汤加", "特立尼达和多巴哥", "突尼斯",
                    "土耳其", "土库曼斯坦", "特克斯科斯群岛", "图瓦卢", "乌干达", "乌克兰", "阿联酋", "英国", "美国", "美国本土外小岛屿", "乌拉圭", "乌兹别克斯坦",
                    "瓦努阿图", "梵蒂冈", "委内瑞拉", "越南", "英属维尔京群岛", "美属维尔京群岛", "瓦利斯和富图纳", "西撒哈拉", "也门", "南斯拉夫", "赞比亚", "津巴布韦",
                    "塞尔维亚"]
    householder_relation_list = ["户主", "配偶", "儿子", "女儿", "父亲", "母亲", "外祖父", "外祖母", "祖父", "祖母", "儿媳", "女婿", "孙子", "孙女",
                                 "外孙女", "兄弟", "姐妹", "其他亲属关系", "无亲属关系", "未知", "外孙子", "婆婆", "岳父", "岳母", "公公", nan]
    dict1 = {"证件类型": id_type_list,
             "政治面貌": politic_status_list,
             "文化程度": edu_list,
             "人员类型": people_type_list,
             "人员状态": people_status_list,
             "居住状态": live_status_list,
             "婚姻状况": marry_list,
             "性别": sex_list,
             "民族": group_list,
             "国家及地区": country_list,
             "是否户主/与户主关系": householder_relation_list}
    df['规则校验列'] = ''
    for key, value in dict1.items():
        df.loc[~df[key].isin(value), ['规则校验列']] += ',' + key + '列不符合输入规则'


def add_build_and_room(df):
    df['楼幢'] = df['楼幢'].apply(lambda x: str(x) + '幢' if '幢' not in str(x) else x)
    df['房间号'] = df['房间号'].apply(lambda x: str(x) + '室' if '室' not in str(x) else x)


def verify_document_number(df):
    df['证件号码校验'] = df['证件号码'].map(lambda x: IdentityCard.is_id_card(x))


def verify_contact_col(df: pd.DataFrame):
    def is_contact(contact):
        if len(str(contact)) != 11 & len(str(contact)) != 0:
            return ',联系方式有误'
        else:
            return ''

    df['联系方式长度校验列'] = df['联系方式'].apply(lambda x: is_contact(x))
    result = Counter(df['联系方式'])
    result_filter = {key: val for key, val in result.items() if val >= 10}
    # print(result_filter)

    if result_filter:
        list_phone = []
        for key, value in result_filter.items():
            temp = key
            list_phone.append(temp)
        list_phone.remove(nan)
        df['联系方式重复数校验列'] = ''
        df.loc[df['联系方式'].isin(list_phone), ['联系方式重复数校验列']] += ',' + '联系方式重复10次及以上'


def verify_room(df_detail: pd.DataFrame, df_room: pd.DataFrame):
    df_detail['房屋校验列'] = ''
    df_detail.loc[~df_detail['小区'].isin(df_room['小区']), ['房屋校验列']] += ',系统中无此小区'
    df_detail['小区楼幢房间号'] = df_detail['小区'].map(str) + df_detail['楼幢'].map(str) + df_detail['房间号'].map(str)
    df_room['小区楼幢房间号'] = df_room['小区'].map(str) + df_room['楼幢'].map(str) + df_room['房间号'].map(str)
    df_detail.loc[~df_detail['小区楼幢房间号'].isin(df_room['小区楼幢房间号']), ['房屋校验列']] += ',系统中无此房间号'
    # df_detail.drop(['小区楼幢房间号'], axis=1, inplace=True)


def verify_num_of_resident(df: pd.DataFrame):
    df['居住人数校验列'] = ''
    house_count = Counter(df['小区楼幢房间号'])
    house_count_filter = {key: val for key, val in house_count.items() if val > 10}
    if house_count_filter:
        list_house = []
        for key, value in house_count_filter.items():
            temp = key
            list_house.append(temp)
        df.loc[df['小区楼幢房间号'].isin(list_house), '居住人数校验列'] += ',同一房号下居住人数大于10人'
    df.drop(['小区楼幢房间号'], axis=1, inplace=True)


def verify_one_room_duplicate_people(df: pd.DataFrame):
    df['重复姓名证件号码校验列'] = ''
    list_count = ['姓名', '证件号码']
    for i in list_count:
        df['小区楼幢房间号' + i] = \
            df['小区'].map(str) + df['楼幢'].map(str) + df['房间号'].map(str) + df[i].map(str)
        count_name = Counter(df['小区楼幢房间号' + i])
        count_name_filter = {key: val for key, val in count_name.items() if val > 1}
        if count_name_filter:
            list_name = []
            for key, value in count_name_filter.items():
                temp = key
                list_name.append(temp)
            df.loc[df['小区楼幢房间号' + i].isin(list_name), ['重复姓名证件号码校验列']] += ',同一房号下' + i + '重复'
            df.drop(['小区楼幢房间号' + i], axis=1, inplace=True)
        else:
            df.drop(['小区楼幢房间号' + i], axis=1, inplace=True)


def verify_is_exist_in_db(df_detail: pd.DataFrame, df_db: pd.DataFrame):
    df_detail['重复数据校验列'] = ''
    df_detail['小区楼幢房间号证件号码'] = \
        df_detail['小区'].map(str) + df_detail['楼幢'].map(str) + df_detail['房间号'].map(str) + df_detail['证件号码'].map(str)
    df_db['小区楼幢房间号证件号码'] = \
        df_db['小区'].map(str) + df_db['楼幢'].map(str) + df_db['房间号'].map(str) + df_db['证件号码'].map(str)
    df_detail.loc[df_detail['小区楼幢房间号证件号码'].isin(df_db['小区楼幢房间号证件号码']), ['重复数据校验列']] += ',已在知社区存在'
    df_detail.drop(['小区楼幢房间号证件号码'], axis=1, inplace=True)


def combine_err_col(df: pd.DataFrame):
    df['备注'] = ''
    col_list = [x for x in df.columns if '校验列' in x]
    for col in col_list:
        df['备注'] += df[col].map(str)
    df['备注'] = df['备注'].apply(lambda x: x[1:] if len(x) != 0 else x)
    for i in col_list:
        # print(i)
        df.drop(i, axis=1, inplace=True)


def main():
    print('选择需校验的文件')
    path = get_path('选择需要校验的文件')
    df = pd.ExcelFile(path)
    verify_sheet_name(df)  # 校验sheet表
    df_detail = df.parse('明细', dtype={'联系方式': str, '证件号码': str})
    df_room = df.parse('房屋信息')
    df_db = df.parse('人房信息')
    print('校验必填列是否存在')
    verify_required_exist(df_detail, df_room, df_db)  # 校验列是否缺失
    print('校验必填项是否为空')
    verify_required_empty(df_detail, df_room, df_db)  # 校验必填项是否为空
    print('校验是否符合下拉选项')
    verify_input_rule(df_detail)  # 校验输入规则
    print('校验是否有重复数据')
    drop_duplicate_data(df_detail)  # 去除重复数据
    print('证件号码统一转大写')
    col_upper(df_detail)  # 证件号码转大写
    col_upper(df_db)
    print('楼幢、房号列补充幢和室')
    add_build_and_room(df_detail)  # 补充幢和室
    print('校验证件号码')
    verify_document_number(df_detail)  # 校验证件号码
    print('校验联系方式位数、是否重复10次及以上')
    verify_contact_col(df_detail)  # 校验联系方式列 1.联系方式位数 2.联系方式是否重复10次及以上
    print('校验明细中房屋信息是否存在')
    verify_room(df_detail, df_room)  # 校验明细中房屋信息是否存在
    print('校验明细中同一房间下居住人数是否大于10人')
    verify_num_of_resident(df_detail)  # 校验明细中同一房间下居住人数是否大于10人
    print('校验统一房间下是否有重复姓名、证件号码')
    verify_one_room_duplicate_people(df_detail)  # 校验明细中同一房号下重复姓名、证件号码
    print('校验是否已录入知社区')
    verify_is_exist_in_db(df_detail, df_db)  # 校验明细中人房信息是否在知社区中存在
    combine_err_col(df_detail)  # 检验列合并
    new_path = path.replace(os.path.basename(os.path.normpath(path)), "校验后.xlsx")
    df_detail.to_excel(new_path, index=False)
    print(f'校验完成，校验后文件保存在{new_path}')
    time.sleep(20)


if __name__ == '__main__':
    if 'win' in sys.platform:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 解决gui模糊的问题
    main()
