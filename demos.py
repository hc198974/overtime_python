# -*- coding: utf-8 -*-
import calendar
import datetime
import tkinter
import tkinter.simpledialog
import requests
from lxml import etree
import win32com.client
import time
import functools


def run_time(fn):  # 用于测试方法运行时间的装饰器
    @functools.wraps(fn)
    def wrapper(*args, **kw):
        start = time.perf_counter()
        res = fn(*args, **kw)
        print('%s 运行了 %f 秒' % (fn, time.perf_counter() - start))
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
        excel.Quit()


class Ccal:
    def __init__(self,year,month):
        self.year = year
        self.month = month
        
    def get_current_month_holidays(self):
        url = f"https://timor.tech/api/holiday/year/{self.year}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"}
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()  # 检查请求是否成功
            data = response.json()
            holidays = []
            for date, info in data["holiday"].items():
                # 检查日期是否为当前月份
                if int(date.split("-")[0]) == self.month and info["holiday"]:
                    holidays.append(date)
            return holidays
        except requests.RequestException as e:
            print(f"请求出错: {e}")
        except (KeyError, ValueError) as e:
            print(f"解析数据出错: {e}")
        return []
