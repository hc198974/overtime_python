# -*- coding: utf-8 -*-
import calendar
import datetime
import itertools
import tkinter
import tkinter.simpledialog
import requests
from lxml import etree
from openpyxl import load_workbook
import win32com.client
import time
import functools
import pandas as pd
import numpy as np


def run_time(fn):  # 用于测试方法运行时间的装饰器
    @functools.wraps(fn)
    def wrapper(*args, **kw):
        start = time.time()
        res = fn(*args, **kw)
        print('%s 运行了 %f 秒' % (textname, time.time() - start))
        return res
    return wrapper


class Crili(object):
    """
    万年日历接口数据抓取
    Params:year 四位数年份字符串
    """

    def __init__(self, year, month):
        self.year = year
        self.month = month

    def parseHTML(self):
        """页面解析"""
        global weekday
        url = "https://wannianrili.bmcx.com/ajax/"
        s = requests.session()
        headers = {
            "Host": "wannianrili.bmcx.com",
            "Connection": "keep-alive",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36",
            "Accept": "*/*",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Dest": "empty",
            "Referer": "https://wannianrili.51240.com/",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        }
        result = {}

        c = calendar.monthrange(self.year, self.month)[1]
        s = requests.session()
        payload = {"q": str(self.year) + "-" + str(self.month)}
        response = s.get(url, headers=headers, params=payload)
        element = etree.HTML(response.text)
        html = element.xpath('//div[@class="wnrl_riqi"]')

        # 获取节点属性
        for i in range(c):
            item = html[i].xpath("./a")[0].attrib
            if item["id"] == "wnrl_riqi_id_" + str(i):
                if "class" in item:
                    temp = datetime.datetime(self.year, self.month, i + 1)
                    if item["class"] == "wnrl_riqi_xiu":
                        weekday = 3
                    elif item["class"] == "wnrl_riqi_mo":
                        weekday = 2
                    elif item["class"] == "wnrl_riqi_ban":
                        weekday = 1.5
                else:
                    temp = datetime.datetime(self.year, self.month, i + 1)
                    if temp.weekday() > 4:
                        weekday = 2
                    else:
                        weekday = 1.5

                result[temp.strftime("%Y%m%d")] = weekday
        # 字典转换为pandas.Series
        result = pd.DataFrame({'日报日期': result.keys(), '节假日': result.values()})
        return result


class Cwindow(object):
    def __init__(self):
        self.month = datetime.datetime.now().month - 1

    def set_win_center(self, root, curWidth="", curHight=""):
        """
        设置窗口大小，并居中显示
        param root:主窗体实例
        param curWidth:窗口宽度，非必填，默认200
        return:无
        """
        if not curWidth:
            """获取窗口宽度，默认200"""
            curWidth = root.winfo_width()
        if not curHight:
            """获取窗口高度，默认200"""
            curHight = root.winfo_height()

        # 获取屏幕宽度和高度
        scn_w, scn_h = root.maxsize()

        # 计算中心坐标
        cen_x = (scn_w - curWidth) / 2
        cen_y = (scn_h - curHight) / 2

        # 设置窗口初始大小和位置
        size_xy = "%dx%d+%d+%d" % (curWidth, curHight, cen_x, cen_y)
        root.geometry(size_xy)

    def askName(self):
        # 获取字符串（标题，提示，初始值）
        name = tkinter.simpledialog.askstring(
            title="获取信息", prompt="请输入姓名：", initialvalue="韩超"
        )
        self.name = name

    def askMonth(self):
        month = tkinter.simpledialog.askinteger(
            title="获取月份",
            prompt="请输入月份",
            initialvalue=datetime.datetime.now().month - 1,
        )
        self.month = month

    def dealSheet(self):
        # 对汇总表数据进行处理
        Cmacro().dealData()

    def shutDown(self):
        root.destroy()

    def createWindow(self):
        global root
        # 创建主窗口
        root = tkinter.Tk()
        # 设置窗口大小
        root.resizable(False, False)
        root.title("加班")
        root.update()
        self.set_win_center(root, 300, 150)
        # 添加按钮
        # btn1 = tkinter.Button(root, text='获取用户名', command=self.askName)
        # btn1.pack(expand='yes')
        btn2 = tkinter.Button(root, text="获取月份", command=self.askMonth)
        btn2.pack(expand="yes")
        btn4 = tkinter.Button(root, text="处理数据", command=self.dealSheet)
        btn4.pack(expand="yes")
        btn3 = tkinter.Button(root, text="开始计算", command=self.shutDown)
        btn3.pack(expand="yes")
        # 加入消息循环
        root.mainloop()


class Cmacro:
    def __init__(self) -> None:
        self.path = (
            r"c:\Users\Administrator\Documents\GitHub\overtime_python\原始数据.xlsm"
        )

    def dealData(self):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(self.path)
        print("START")
        excel.Application.Run("deleteRow")
        wb.SaveAs(
            r"c:\Users\Administrator\Documents\GitHub\overtime_python\计算结果.xlsx",
            FileFormat=51,
            ConflictResolution=2,
        )
        wb.Close()
        print("END")


def custom_gettime(row):
    result = 0
    sb = row['上班时间']
    xb = row['下班时间']
    temp17 = datetime.datetime.strptime("17:30", "%H:%M")
    temp18 = datetime.datetime.strptime("18:00", "%H:%M")
    temp12 = datetime.datetime.strptime("12:00", "%H:%M")
    temp13 = datetime.datetime.strptime("13:00", "%H:%M")
    temp8 = datetime.datetime.strptime("8:00", "%H:%M")
    if not pd.isna(sb) and not pd.isna(xb):
        if type(sb) == str:
            sb = datetime.datetime.strptime(sb, "%H:%M")
        elif type(sb) == datetime.time:
            sb = datetime.datetime.strptime(sb.strftime("%H:%M"))

        if type(xb) == str:
            xb = datetime.datetime.strptime(xb, "%H:%M")
        elif type(xb) == datetime.time:
            xb = datetime.datetime.strptime(xb.strftime("%H:%M"))
        if row['节假日'] == 1.5:
            if xb >= sb and sb <= temp17:
                if xb >= temp18:
                    result = round(((xb-temp17).seconds)/3600, 2)
            else:
                result = 0
        else:
            if xb >= sb:
                if sb < temp8:
                    sb = temp8
                if sb > temp12 and sb < temp13:
                    sb = temp13
                if xb > temp12 and xb < temp13:
                    xb = temp13

                if xb > sb:
                    result = round(((xb-sb).seconds)/3600, 2)
            else:
                result = 0
    return result


def getgroup(df, group3, group2, group1):
    jiaban = []
    remainer = 36
    if not group3.empty:
        # 先默认3倍加班费超不过36小时
        jiaban.append(group3.index)
        remainer = remainer - group3['时长'].sum()

    if not group2.empty:
        if remainer >= group2['时长'].sum():
            jiaban.append(group2.index)
            remainer = remainer - group2['时长'].sum()
        else:
            coms = []
            for i in range(len(group2.loc[group2['时长'] != 0, ['时长']])):
                combinations = list(
                    itertools.combinations(list(group2[group2['时长'] != 0].index), i+1))
                coms.append(combinations)

            max = 0
            total = 0
            # coms范例：[[(247,), (255,), (261,)], [(247, 255), (247, 261), (255, 261)], [(247, 255, 261)]]
            for c in coms:
                for x in c:
                    for z in x:
                        total = total+df.loc[z, ['时长']].values
                    if total <= remainer:
                        if total > max:
                            max = total
                            df_max = x

            jiaban.append(df_max)
            # group2.loc[~group2['日报日期'].isin(df_max), group2['时长'] != 0,['加班或串休']] = '转串休'
            remainer = remainer - max

    if not group1.empty:
        if remainer >= group1['时长'].sum():
            group1.loc[group1['时长'] != 0, ['加班或串休']] = 1
            remainer = remainer - group1['时长'].sum()
        else:
            coms = []
            for i in range(len(group1.loc[group1['时长'] != 0, ['时长']])):
                combinations = list(
                    itertools.combinations(list(group1[group1['时长'] != 0].index), i+1))
                coms.append(combinations)

            max = 0
            total = 0
            # coms范例：[[(247,), (255,), (261,)], [(247, 255), (247, 261), (255, 261)], [(247, 255, 261)]]
            for c in coms:
                for x in c:
                    total = sum(
                        list(map(lambda z: df.loc[z, ['时长']].values, x)))
                    if total <= remainer:
                        if total > max:
                            max = total
                            df_max = x
            jiaban.append(df_max)
            remainer = remainer - max
    print(jiaban)

def custom_getgroup(df, group):
    if group[group['时长'] != 0].empty:
        print('空')
    elif group['时长'].sum() <= 36:
        group.loc[group['时长'] != 0, ['加班或串休']] = 1
    else:
        group3 = group[group['节假日'] == 3]
        group2 = group[group['节假日'] == 2]
        group1 = group[group['节假日'] == 1.5]
        getgroup(df, group3, group2, group1)


def main(result):
    # 读取Excel文件，默认第一个表《汇总表》
    df = pd.read_excel('计算结果.xlsx')
    df['日报日期'] = df['日报日期'].dt.strftime('%Y%m%d')
    df.drop('节假日', axis=1, inplace=True)
    df = df.merge(result)
    df['时长'] = df.apply(lambda row: custom_gettime(row), axis=1)
    # print(df)
    # # 分组计算
    grouped = df.groupby(['姓名'])
    for name, group in grouped:
        custom_getgroup(df, group)


if __name__ == "__main__":
    # cw = Cwindow()
    # cw.createWindow()
    result = Crili(2025, 1).parseHTML()
    start = time.perf_counter()
    main(result)
    end = time.perf_counter()
    print("运行时间：", end - start)
