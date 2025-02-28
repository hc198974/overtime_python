# -*- coding: utf-8 -*-
import calendar
import datetime
import itertools
import requests
from lxml import etree
import time
import functools
import pandas as pd
import numpy as np


def run_time(fn):  # 用于测试方法运行时间的装饰器
    @functools.wraps(fn)
    def wrapper(*args, **kw):
        start = time.time()
        res = fn(*args, **kw)
        print('%s 运行了 %f 秒' % (fn, time.time() - start))
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


@run_time
def overtime_cal(datas, result):
    """加班计算"""
    for i in range(len(datas)//len(result)):
        chunk = datas[i*len(result):(i*len(result))+(len(result)-1)]
        chunk = dealdata(chunk, result)
        chunk = maximize_value(chunk)
        # modified_chunk = np.append(chunk, dealdata(chunk, result), axis=1)


def maximize_value(result, max_sum=36000):

    if sum([row[7] for row in result]) <= max_sum/1000:
        for row in result:
            if row[7] > 0:
                row[6] = '转加班'
        return result
    else:

        weights = [int(round(row[7], 2)*1000) for row in result]
        values = [int(round(row[7], 2) * row[5]*1000)
                  for row in result]
        n = len(weights)
        # 创建动态规划表
        dp = [[0] * (max_sum + 1) for _ in range(n + 1)]

        # 填充动态规划表
        for i in range(1, n + 1):
            for w in range(max_sum + 1):
                if weights[i-1] <= w:
                    dp[i][w] = max(dp[i-1][w], dp[i-1]
                                   [w-weights[i-1]] + values[i-1])
                else:
                    dp[i][w] = dp[i-1][w]

        # 回溯找到选择的物品
        w = max_sum
        items = []
        for i in range(n, 0, -1):
            if dp[i][w] != dp[i-1][w]:
                items.append(list(result[i-1]))
                w -= weights[i-1]

        # 检查某一行是否在 selected_items 中
        for row in result:
            row1 = list(row)
            if row1 in items:
                if row[7] > 0:
                    row[6] = '转加班'
            else:
                if row[7] > 0:
                    row[6] = '转串休'
        return result


def dealdata(chunk, result):
    def calculate_time(sb, xb, isweekend):
        temp17 = datetime.datetime.strptime("17:30", "%H:%M").time()
        temp18 = datetime.datetime.strptime("18:00", "%H:%M").time()
        temp12 = datetime.datetime.strptime("12:00", "%H:%M").time()
        temp13 = datetime.datetime.strptime("13:00", "%H:%M").time()
        temp8 = datetime.datetime.strptime("8:00", "%H:%M").time()

        if pd.isna(sb) or pd.isna(xb):
            return 0
        if isinstance(sb, str):
            sb = datetime.datetime.strptime(sb, "%H:%M").time()
        if isinstance(xb, str):
            xb = datetime.datetime.strptime(xb, "%H:%M").time()
        if isweekend == 1.5:
            if xb >= sb and sb <= temp17:
                if xb >= temp18:
                    return round((datetime.datetime.combine(datetime.date.today(), xb) - datetime.datetime.combine(datetime.date.today(), temp17)).seconds / 3600, 2)
            return 0
        else:
            delta = round((datetime.datetime.combine(datetime.date.today(
            ), xb) - datetime.datetime.combine(datetime.date.today(), sb)).seconds/3600, 2)
            if delta > 0.5:
                if sb < temp8:
                    sb = temp8
                if sb > temp12 and sb < temp13:
                    sb = temp13
                if xb > temp12 and xb < temp13:
                    xb = temp13
                # 计算加班时间
                delta = round((datetime.datetime.combine(datetime.date.today(
                ), xb) - datetime.datetime.combine(datetime.date.today(), sb)).seconds/3600, 2)
                if xb >= temp13 and sb <= temp12:
                    return delta-1.5
                else:
                    return delta-0.5
            return 0
    for i in chunk:
        # 先导入节假日
        if i[1] in result:
            i[5] = result[i[1]]
        # 第二步计算加班小时数
        temp = calculate_time(i[2], i[3], i[5])
        if temp < 0:
            temp = 0
        i[7] = temp

    # 第三步计算加班金额
    return chunk


if __name__ == "__main__":
    start = time.perf_counter()
    # 获得工作日和节假日
    result = Crili(2025, 1).parseHTML()
    df = pd.read_excel('计算结果.xlsx')
    df['日报日期'] = df['日报日期'].dt.strftime('%Y%m%d')
    datas = df.to_numpy()
    sorted_indices = np.argsort(datas[:, 0])  # 获取排序索引
    datas = datas[sorted_indices]  # 按索引重新排列数组
    overtime_cal(datas, result)
    end = time.perf_counter()
    print("运行时间：", end - start)
