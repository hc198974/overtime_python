# -*- coding: utf-8 -*-
import calendar
import datetime
import itertools
import requests
from lxml import etree
from openpyxl import load_workbook
import time
import functools
import pandas as pd
import numpy as np

# 全局变量
testname = ""


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
        return result


class Count(object):
    def __init__(self, names, month, result, wb):
        # 节假日接口(工作日对应结果为 0, 休息日对应结果为 1, 节假日对应的结果为 2 )
        # server_url = "http://www.easybots.cn/api/holiday.php?d="
        self.server_url = "http://tool.bitefu.net/jiari/?d="
        self.wb = wb
        self.ws = self.wb["汇总表"]
        self.ws2 = self.wb["中干"]
        self.names = names
        self.name = ""
        self.month = month
        self.dict = {}
        self.dictall = {}
        self.weekday = {}
        self.workday = {}
        self.holiday = {}
        self.cash = {}
        self.hour = 0
        self.result = result

    def getUrl(self):
        try:
            for m in self.result:
                if self.result[m] == 1.5:
                    self.workday[m] = 0
                elif self.result[m] == 2:
                    self.weekday[m] = 1
                elif self.result[m] == 3:
                    self.holiday[m] = 2
        except ConnectionResetError as e:
            print("远程主机发生错误" + e)

    # 调整表里的加班小时数
    def changeHour(self):
        temp17 = datetime.datetime.strptime("17:30", "%H:%M")
        temp18 = datetime.datetime.strptime("18:00", "%H:%M")
        temp12 = datetime.datetime.strptime("12:00", "%H:%M")
        temp13 = datetime.datetime.strptime("13:00", "%H:%M")
        temp8 = datetime.datetime.strptime("8:00", "%H:%M")
        self.getUrl()

        for name in self.names:
            dict = {}
            for x in self.ws.rows:
                if x[4].value == self.month:
                    if x[0].value == name.value:
                        temp = x[1].value.strftime("%Y%m%d")
                        time1 = x[2].value
                        time2 = x[3].value
                        if (
                            time1 != ""
                            and time2 != ""
                            and time1 is not None
                            and time2 is not None
                        ):
                            if type(time1) == str:
                                time1 = datetime.datetime.strptime(
                                    time1, "%H:%M")
                            elif type(time1) == datetime.time:
                                time1 = datetime.datetime.strptime(
                                    (time1.strftime("%H:%M")), "%H:%M"
                                )  # 先把datetime.time格式转换为str再转换为datetime.datetime
                            if type(time2) == str:
                                time2 = datetime.datetime.strptime(
                                    time2, "%H:%M")
                            elif type(time2) == datetime.time:
                                time2 = datetime.datetime.strptime(
                                    (time2.strftime("%H:%M")), "%H:%M"
                                )

                            if time2 > time1:
                                # 工作日
                                if temp in self.workday:
                                    if time2 > temp18:
                                        self.hour = (time2 - temp17).seconds

                                if self.hour > 0:
                                    x[7].value = round(self.hour / 3600, 2)
                                    s = x[1].value.strftime("%Y%m%d")
                                    dict[s] = x[7].value
                                    x[5].value = "工作日"
                                    self.hour = 0
                                else:
                                    x[7].value = 0
                                    x[5].value = "工作日"
                                    self.hour = 0

                                # 周末和节假日
                                if temp in self.weekday or temp in self.holiday:
                                    if time1 > temp8:
                                        if temp12 < time1 < temp13:
                                            time1 = temp12
                                    else:
                                        time1 = temp8

                                    if temp12 < time2 < temp13:
                                        time2 = temp13
                                    else:
                                        pass

                                    if time2 <= temp12:
                                        self.hour = (
                                            time2 - time1 -
                                            datetime.timedelta(hours=0.5)
                                        )
                                    if time2 >= temp13:
                                        if time1 <= temp12:
                                            self.hour = (
                                                time2
                                                - time1
                                                - datetime.timedelta(hours=1.5)
                                            )
                                        else:
                                            self.hour = (
                                                time2
                                                - time1
                                                - datetime.timedelta(hours=0.5)
                                            )

                                    if self.hour.days == 0:
                                        x[7].value = round(
                                            self.hour.seconds / 3600, 2)
                                        s = x[1].value.strftime("%Y%m%d")
                                        dict[s] = x[7].value
                                        x[5].value = "节假日"
                                        self.hour = 0
                                    else:
                                        x[7].value = 0
                                        x[5].value = "节假日"
                                        self.hour = 0
                self.dictall[name.value] = dict

    def setContents(self):
        # 在中干表写入是转加班费还是串休

        for x in self.ws.rows:
            if not x[7].value is None:
                if x[0].value == self.name.value and x[7].value > 0:
                    if x[1].value.strftime("%Y%m%d") in self.cash.keys():
                        x[6].value = "转加班费"
                    else:
                        x[6].value = "转串休"

        rng = self.ws2["C2":"AG2"]
        for x in rng:
            for y in x:
                for z in self.dict:
                    if y.value.strftime("%Y%m%d") == z:
                        self.ws2.cell(row=self.name.row, column=y.column).value = self.dict[
                            z
                        ]
        self.ws2.cell(row=self.name.row, column=34).value = sum(
            list(self.dict.values()))
        self.ws2.cell(row=self.name.row, column=35).value = sum(
            list(self.dict.values()))-sum(list(self.cash.values()))

    @run_time
    def count2name(self):
        global textname
        textname = self.name.value
        self.cash.clear()
        self.dict = self.dictall.get(self.name.value)
        dict_1, dict_2, dict_3 = {}, {}, {}
        for x in self.dict:
            if self.result[x] == 1.5:
                dict_1.update({x: self.dict[x]})
            elif self.result[x] == 2:
                dict_2.update({x: self.dict[x]})
            elif self.result[x] == 3:
                dict_3.update({x: self.dict[x]})

        remainder = 36
        if sum(list(self.dict.values())) > 36:
            for p in [dict_3, dict_2, dict_1]:
                if len(p) > 0:
                    combine = []
                    for r in range(1, len(p) + 1):
                        combinations = list(itertools.combinations(p, r))
                        for x in combinations:
                            combine.append(x)

                    temp = {}
                    smax = 0
                    total = 0
                    for m in combine:
                        for n in m:
                            total += self.dict[n]

                        if total <= remainder:
                            if smax < total:
                                smax = total
                                temp.clear()
                                for y in m:
                                    temp.update({y: self.dict[y]})
                                total = 0
                            else:
                                total = 0
                        else:
                            total = 0

                    self.cash.update(temp)
                    temp.clear
                    remainder = remainder - sum(list(self.cash.values()))

        else:
            self.cash = self.dict.copy()

        self.setContents()

    def jiSuan(self):
        # 获得URL
        self.changeHour()
        for row in self.ws2.iter_rows(min_row=3, max_row=self.ws2.max_row, min_col=3, max_col=35):
            for cell in row:
                cell.value = None

        for name in self.names:
            # 数据量不大，使用多进程开销大，多线程容易出现错误
            self.name = name
            self.count2name()


if __name__ == "__main__":
    start = time.perf_counter()
    # 获得工作日和节假日
    result = Crili(2025, 1).parseHTML()
    df = pd.read_excel('计算结果.xlsx')
    df2=pd.read_excel('计算结果.xlsx',sheet_name='中干')    
    df['日报日期'] = df['日报日期'].dt.strftime('%Y%m%d')
    df.drop('节假日', axis=1, inplace=True)
    # df = df.merge(result)
    n1=df.to_numpy()
    n2=df2.iloc[:,1].to_numpy()
    print(n1)
    end = time.perf_counter()
    print("运行时间：", end - start)
