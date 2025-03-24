#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# pythonproject -> 一件事信息爬取
# @IDE      ：PyCharm
# @Author   ：Mob
# @Date     ：2023/12/15 13:51
# @Software : win10 python3.9
import time

import requests
import pandas as pd


class Yjs:
    def __init__(self):
        self.cookies = {
            '_font_size_ratio_': '1.0',
            'sid': '8d6e9e3baeca7902f49ed46d6a31d394',
            'access_token_expiresin': '3600',
            'access_token': '8d6e9e3baeca7902f49ed46d6a31d394',
            'refresh_token': 'b715e58aba0bf7c4e655f34877ab6e95',
        }

        self.headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            # 'Cookie': '_font_size_ratio_=1.0; sid=8d6e9e3baeca7902f49ed46d6a31d394; access_token_expiresin=3600; access_token=8d6e9e3baeca7902f49ed46d6a31d394; refresh_token=b715e58aba0bf7c4e655f34877ab6e95',
            'Origin': 'http://172.26.199.170:81',
            'Referer': 'http://172.26.199.170:81/epoint-zwfwbase-project-web/jsproject/onethingyicibantj/one_thing/index?isopen=0',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36',
            'User-Token': 'userSign=953E92E41BADAC1D82B2ECB2A23EF0999FFFB632079ADF7B566425ADB44876E6,reqTime=1730092436292,deviceId=f9db48942e46ea81adfea2c5e9404981,title=%E4%B8%80%E4%BB%B6%E4%BA%8B%E5%8A%9E%E4%BB%B6%E7%BB%9F%E8%AE%A1',
            'X-Requested-With': 'XMLHttpRequest',
        }

        self.params = {
            'isopen': '0',
            'isCommondto': 'true',
        }

    def get_data(self, area_code, area_name):
        data = {
            'commonDto': '[{"id":"datagrid","type":"datagrid","action":"getDataGridData","name":"","idField":"id","pageIndex":0,"sortField":"","sortOrder":"","columns":[{"fieldName":"onethingname"},{"fieldName":"scenetype","code":"场景类型"},{"fieldName":"serviceobject","code":"服务对象"},{"fieldName":"rowguid"},{"fieldName":"onethingtype"},{"fieldName":"totalcount"},{"fieldName":"istysl"},{"fieldName":"ouname"},{"fieldName":"promotecitycount"},{"fieldName":"ybjl"},{"fieldName":"selfmadecitycount"},{"fieldName":"selfmadecity"}],"pageSize":-1,"url":"getDataGridData","data":[]},{"id":"areaCode","bind":"area","type":"combobox","action":"getAreacode","name":"","pinyinField":"tag","columns":[],"textField":"text","valueField":"id","value":"' + area_code + '","text":"' + area_name + '"},{"id":"dateFirst","bind":"time","type":"combobox","action":"getTimeType","name":"","pinyinField":"tag","columns":[],"textField":"text","valueField":"id","value":"year","text":"本年"},{"id":"sceneName","bind":"onethingname","type":"textbox","action":"","name":"","value":"","text":""},{"id":"source","bind":"source","type":"combobox","action":"","name":"source","pinyinField":"tag","columns":[],"textField":"text","valueField":"id","value":"all","text":"全部"},{"id":"dataexport","type":"dataexport","action":"getExportModel","name":"","mapClass":"com.epoint.basic.faces.export.DataExport","exportAction":"","text":"导出"},{"id":"_common_hidden_viewdata","type":"hidden","value":"{\\"epoint_user_loginid\\":\\"953E92E41BADAC1D82B2ECB2A23EF0999FFFB632079ADF7B566425ADB44876E6\\",\\"pageurl\\":\\"923d724527d63404952bd04250b22d8a50174ba998e101f90c76c16a65e99127972981051e244c59cad7c38ba698cc84a32918e9f30de9ed88ac72d86cbf84a5\\"}"}]',
            'cmdParams': '{"pageUrl":"http://172.26.199.170:81/epoint-zwfwbase-project-web/jsproject/onethingyicibantj/one_thing/index?isopen=0"}',
            'replaynoticeid': '26122AC6-4A23-4B13-B986-90BD73ADEBE9',
            'reqtime': '1730092436247',
            'paramsign': '3B399521DC7E277CD6D65A9719BA15FD4C56C70577F8242E2A3284F1CF24963B',
            'dtosign': '4E5BE8797066524E5C38E3B3F4267B92E254685E50C4E1B51473C641C771B619',
        }

        response = requests.post(
            'http://172.26.199.170:81/epoint-zwfwbase-project-web/rest/anascenebuildingsituationaction/page_Refresh',
            params=self.params,
            cookies=self.cookies,
            headers=self.headers,
            data=data,
            verify=False,
        )
        print(response)
        print(response.json())
        df = pd.DataFrame(response.json()['controls'][0]['data'])
        return df


if __name__ == '__main__':
    areas = [
        {'area_name': '南京', 'area_code': '3201'},
        {'area_name': '无锡', 'area_code': '3202'},
        {'area_name': '徐州', 'area_code': '3203'},
        {'area_name': '常州', 'area_code': '3204'},
        {'area_name': '苏州', 'area_code': '3205'},
        {'area_name': '南通', 'area_code': '3206'},
        {'area_name': '连云港', 'area_code': '3207'},
        {'area_name': '淮安', 'area_code': '3208'},
        {'area_name': '盐城', 'area_code': '3209'},
        {'area_name': '扬州', 'area_code': '3210'},
        {'area_name': '镇江', 'area_code': '3211'},
        {'area_name': '泰州', 'area_code': '3212'},
        {'area_name': '宿迁', 'area_code': '3213'}
    ]
    dfs = []
    for area in areas:
        time.sleep(2)
        df = Yjs().get_data(area_name=area['area_name'], area_code=area['area_code'])
        column_mapping = {
            "totalcount": str(area['area_name']) + "办件总量（件）",
            "scenetype": "事项类型",
            "onethingname": "应用场景",
            "serviceobject": "办理主体",
            "ybjl": str(area['area_name']) + "办结率"

        }
        columns_to_keep = [col for col in df.columns if col in column_mapping.keys()]
        df = df[columns_to_keep]
        print(df)
        df = df.rename(columns=column_mapping)
        column_order = ['事项类型', '应用场景', '办理主体', str(area['area_name']) + "办件总量（件）", str(area['area_name']) + "办结率"]
        df = df.reindex(columns=column_order)
        dfs.append(df)
    for i in dfs[1:]:
        i.drop(columns=['事项类型', '应用场景', '办理主体'], inplace=True)
    df_merged = pd.concat(dfs, axis=1)
    columns_bjl = df_merged.columns[df_merged.columns.str.contains('办件总量（件）')]
    columns_rate = df_merged.columns[df_merged.columns.str.contains('办结率')]
    df_bjl = df_merged.loc[:, ['事项类型', '应用场景', '办理主体'] + list(columns_bjl)]
    df_bjl.columns = df_bjl.columns.str.replace('办件总量（件）', '')
    df_bjl = df_bjl[
        ['事项类型', '应用场景', '办理主体', '南京', '无锡', '徐州', '常州', '南通', '连云港', '淮安', '盐城', '扬州', '镇江', '泰州', '宿迁', '苏州']]
    df_rate = df_merged.loc[:, ['事项类型', '应用场景', '办理主体'] + list(columns_rate)]
    df_rate.columns = df_rate.columns.str.replace('办结率', '')
    df_rate = df_rate[[
        '事项类型', '应用场景', '办理主体', '南京', '无锡', '徐州', '常州', '南通', '连云港', '淮安', '盐城', '扬州', '镇江', '泰州', '宿迁', '苏州'
    ]]
    df_bjl.to_excel(r'C:\Users\74752\Desktop\13市办件量.xlsx', index=False)
    df_rate.to_excel(r'C:\Users\74752\Desktop\13市办结率.xlsx', index=False)
