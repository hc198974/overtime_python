# -*- coding: utf-8 -*-
import calendar
import datetime
import itertools
import tkinter
import tkinter.simpledialog
import requests
from lxml import etree
import win32com.client
import time
import functools
import pandas as pd


def run_time(fn):  # 用于测试方法运行时间的装饰器
    @functools.wraps(fn)
    def wrapper(*args, **kw):
        start = time.time()
        res = fn(*args, **kw)
        print('%s 运行了 %f 秒' % (fn.__name__, time.time() - start))
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
        url = "https://wannianrili.bmcx.com/ajax/"
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


def custom_gettime(df):
    temp17 = datetime.datetime.strptime("17:30", "%H:%M").time()
    temp18 = datetime.datetime.strptime("18:00", "%H:%M").time()
    temp12 = datetime.datetime.strptime("12:00", "%H:%M").time()
    temp13 = datetime.datetime.strptime("13:00", "%H:%M").time()
    temp8 = datetime.datetime.strptime("8:00", "%H:%M").time()

    def calculate_time(row):
        sb = row['上班时间']
        xb = row['下班时间']
        if pd.isna(sb) or pd.isna(xb):
            return 0
        if isinstance(sb, str):
            sb = datetime.datetime.strptime(sb, "%H:%M").time()
        if isinstance(xb, str):
            xb = datetime.datetime.strptime(xb, "%H:%M").time()
        if row['节假日'] == 1.5:
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

    df['时长'] = df.apply(calculate_time, axis=1)
    return df


def getgroup(dict, group3, group2, group1):
    jiaban = []
    remainer = 36
    if not group3.empty:
        # 先默认3倍加班费超不过36小时
        jiaban.append(tuple(group3.index))
        remainer = remainer - group3['时长'].sum()

    if not group2.empty:
        if remainer >= group2['时长'].sum():
            jiaban.append(tuple(group2.index))
            remainer = remainer - group2['时长'].sum()
        else:
            coms = []
            for i in range(len(group2.loc[group2['时长'] != 0, ['时长']])):
                combinations = list(
                    itertools.combinations(list(group2[group2['时长'] != 0].index), i+1))
                coms.append(combinations)

            max = 0
            # coms范例：[[(247,), (255,), (261,)], [(247, 255), (247, 261), (255, 261)], [(247, 255, 261)]]
            tuples = [t for sublist in coms for t in sublist]
            for indexes in tuples:
                total = 0
                for z in indexes:
                    total += dict[z]
                if total <= remainer:
                    if total > max:
                        max = total
                        df_max = indexes

            jiaban.append(df_max)
            remainer = remainer - max

    if not group1.empty:
        if remainer >= group1['时长'].sum():
            remainer = remainer - group1['时长'].sum()
        else:
            coms = []
            for i in range(len(group1)):
                combinations = list(
                    itertools.combinations(list(group1.index), i+1))
                coms.append(combinations)
            # coms范例：[[(247,), (255,), (261,)], [(247, 255), (247, 261), (255, 261)], [(247, 255, 261)]]
            max = 0
            tuples = [t for sublist in coms for t in sublist]
            for indexes in tuples:
                total = 0
                for z in indexes:
                    total += dict[z]
                if total <= remainer:
                    if total > max:
                        max = total
                        df_max = indexes
            try:
                jiaban.append(df_max)
            except:
                print(remainer)
    return jiaban


def custom_getgroup(dict, group):
    jiaban = []
    if group[group['时长'] != 0].empty:
        pass
    elif group['时长'].sum() <= 36:
        jiaban.append([group[group['时长'] != 0].index])
    else:
        group3 = group[(group['节假日'] == 3) & (group['时长'] > 0)]
        group2 = group.query('节假日 == 2 & 时长 > 0')
        group1 = group.query('节假日 == 1.5 & 时长 > 0')
        return getgroup(dict, group3, group2, group1)


def generate_summary_table(df):
    # 创建数据透视表，列为日期，index 为姓名，values 为时长求和
    pivot = df.pivot_table(index='姓名', columns='日报日期',
                           values='时长', aggfunc='sum', fill_value=0)
    # 添加合计列
    pivot['合计'] = pivot.sum(axis=1)
    # 重置索引
    summary_df = pivot.reset_index()
    return summary_df


def main(result):
    list = []
    # 读取Excel文件，默认第一个表《汇总表》
    df = pd.read_excel('计算结果.xlsx')
    df['日报日期'] = df['日报日期'].dt.strftime('%Y%m%d')
    df.drop('节假日', axis=1, inplace=True)
    df = df.merge(result)
    df = custom_gettime(df)
    dict = df["时长"].to_dict()
    # # 分组计算
    grouped = df.groupby(['姓名'], sort=True)
    for name, group in grouped:
        result = custom_getgroup(dict, group)
        if not result is None:
            for tuple in result:
                for index in tuple:
                    list.append(index)
    df.loc[df.index.isin(list), '加班或串休'] = 1
    df.loc[(~df.index.isin(list)) & (df["时长"] > 0), '加班或串休'] = 0
    # 新建一个pivot_table
    summary_df = generate_summary_table(df)
    # 多表导出到excel
    start = time.perf_counter()
    with pd.ExcelWriter("site.xlsx") as writer:
        df.to_excel(writer, index=False, sheet_name='明细', engine='openpyxl')
        summary_df.to_excel(writer, index=False,
                            sheet_name='汇总', engine='openpyxl')
    print("时间：", time.perf_counter()-start)


if __name__ == "__main__":
    cw = Cwindow()
    cw.createWindow()
    result = Crili(2025, cw.month).parseHTML()
    start = time.perf_counter()
    main(result)
    end = time.perf_counter()
    print("运行时间：", end - start)
