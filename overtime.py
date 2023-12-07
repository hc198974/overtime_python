# -*- coding: utf-8 -*-
import calendar
import datetime
import tkinter
import tkinter.simpledialog
from itertools import combinations

import requests
from lxml import etree
from openpyxl import load_workbook
import win32com.client


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
        url = 'https://wannianrili.bmcx.com/ajax/'
        s = requests.session()
        headers = {
            'Host': 'wannianrili.bmcx.com',
            'Connection': 'keep-alive',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
            'Accept': '*/*',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
            'Referer': 'https://wannianrili.51240.com/',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        }
        result = {}

        c = calendar.monthrange(self.year, self.month)[1]
        s = requests.session()
        payload = {'q': str(self.year) + '-' + str(self.month)}
        response = s.get(url, headers=headers, params=payload)
        element = etree.HTML(response.text)
        html = element.xpath('//div[@class="wnrl_riqi"]')

        # 获取节点属性
        for i in range(c):
            item = html[i].xpath('./a')[0].attrib
            if item['id'] == 'wnrl_riqi_id_' + str(i):
                if 'class' in item:
                    temp = datetime.datetime(self.year, self.month, i + 1)
                    if item['class'] == 'wnrl_riqi_xiu':
                        weekday = 2
                    elif item['class'] == 'wnrl_riqi_mo':
                        weekday = 1
                else:
                    temp = datetime.datetime(self.year, self.month, i + 1)
                    if temp.weekday() > 4:
                        weekday = 1
                    else:
                        weekday = 0

                result[temp.strftime('%Y%m%d')] = weekday
        return (result)


class Cwindow(object):
    def __init__(self):
        self.month = datetime.datetime.now().month - 1

    def set_win_center(self, root, curWidth='', curHight=''):
        '''
    设置窗口大小，并居中显示
    param root:主窗体实例
    param curWidth:窗口宽度，非必填，默认200
    return:无
    '''
        if not curWidth:
            '''获取窗口宽度，默认200'''
            curWidth = root.winfo_width()
        if not curHight:
            '''获取窗口高度，默认200'''
            curHight = root.winfo_height()

        # 获取屏幕宽度和高度
        scn_w, scn_h = root.maxsize()

        # 计算中心坐标
        cen_x = (scn_w - curWidth) / 2
        cen_y = (scn_h - curHight) / 2

        # 设置窗口初始大小和位置
        size_xy = '%dx%d+%d+%d' % (curWidth, curHight, cen_x, cen_y)
        root.geometry(size_xy)

    def askName(self):
        # 获取字符串（标题，提示，初始值）
        name = tkinter.simpledialog.askstring(
            title='获取信息', prompt='请输入姓名：', initialvalue='韩超')
        self.name = name

    def askMonth(self):
        month = tkinter.simpledialog.askinteger(
            title='获取月份', prompt='请输入月份', initialvalue=datetime.datetime.now().month - 1)
        self.month = month

    def clearSheet(self):
        # 对汇总表数据进行清理
        Cmacro().dealData()

    def shutDown(self):
        root.destroy()

    def createWindow(self):
        global root
        # 创建主窗口
        root = tkinter.Tk()
        # 设置窗口大小
        root.resizable(False, False)
        root.title('加班')
        root.update()
        self.set_win_center(root, 300, 150)
        # 添加按钮
        btn1 = tkinter.Button(root, text='获取用户名', command=self.askName)
        btn1.pack(expand='yes')
        btn2 = tkinter.Button(root, text='获取月份', command=self.askMonth)
        btn2.pack(expand='yes')
        btn4 = tkinter.Button(root, text='清理数据', command=self.clearSheet)
        btn4.pack(expand='yes')
        btn3 = tkinter.Button(root, text='开始计算', command=self.shutDown)
        btn3.pack(expand='yes')
        # 加入消息循环
        root.mainloop()


class Count(object):
    def __init__(self, name, month, result):
        self.fpath = '计算结果.xlsx'
        # 节假日接口(工作日对应结果为 0, 休息日对应结果为 1, 节假日对应的结果为 2 )
        # server_url = "http://www.easybots.cn/api/holiday.php?d="
        self.server_url = "http://tool.bitefu.net/jiari/?d="
        self.wb = load_workbook(filename=self.fpath)
        self.ws = self.wb['汇总表']
        self.ws2 = self.wb['中干']
        self.name = name.value
        self.month = month
        self.dict = {}
        self.weekday = {}
        self.workday = {}
        self.holiday = {}
        self.cash = {}
        self.hour = 0
        self.result = result

    def getUrl(self):
        try:
            for m in self.result:
                if result[m] == 0:
                    self.workday[m] = 0
                elif result[m] == 1:
                    self.weekday[m] = 1
                elif result[m] == 2:
                    self.holiday[m] = 2

        except ConnectionResetError as e:
            print('远程主机发生错误' + e)

    # 调整表里的加班小时数
    def changeHour(self):
        temp17 = datetime.datetime.strptime('17:30', "%H:%M")
        temp18 = datetime.datetime.strptime('18:00', '%H:%M')
        temp12 = datetime.datetime.strptime('12:00', "%H:%M")
        temp13 = datetime.datetime.strptime('13:00', "%H:%M")
        temp8 = datetime.datetime.strptime('8:00', "%H:%M")
        self.getUrl()
        for x in self.ws.rows:
            if x[0].value == self.name:
                if x[4].value == self.month:
                    temp = x[1].value.strftime("%Y%m%d")
                    time1 = x[2].value
                    time2 = x[3].value
                    if time1 != '' and time2 != '' and time1 != None and time2 != None:
                        if time2 > time1:
                            time1 = datetime.datetime.strptime(time1, "%H:%M")
                            time2 = datetime.datetime.strptime(time2, "%H:%M")

                            # 工作日
                            if temp in self.workday:
                                if time2 > temp18:
                                    self.hour = (time2 - temp17).seconds

                            if self.hour > 0:
                                x[7].value = round(self.hour / 3600,2)
                                s = x[1].value.strftime("%Y%m%d")
                                self.dict[s] = x[7].value
                                x[5].value = "工作日"
                                self.hour = 0
                            else:
                                x[7].value = 0
                                x[5].value = "工作日"
                                self.hour = 0

                            # 周末和节假日
                            if temp in self.weekday or temp in self.holiday:
                                if time1 > temp8:
                                    if time1 > temp12 and time1 < temp13:
                                        time1 = temp12
                                else:
                                    time1 = temp8

                                if time2 > temp12 and time2 < temp13:
                                    time2 = temp13
                                else:
                                    pass

                                if time2 <= temp12:
                                    self.hour = (time2 - time1 -
                                                 datetime.timedelta(hours=0.5)).seconds
                                if time2 >= temp13:
                                    if time1 <= temp12:
                                        self.hour = (time2 - time1 -
                                                     datetime.timedelta(hours=1.5)).seconds
                                    else:
                                        self.hour = (time2 - time1 -
                                                     datetime.timedelta(hours=0.5)).seconds

                                if self.hour > 0:
                                    x[7].value = round(self.hour / 3600,2)
                                    s = x[1].value.strftime("%Y%m%d")
                                    self.dict[s] = x[7].value
                                    x[5].value = "节假日"
                                    self.hour = 0
                                else:
                                    x[7].value = 0
                                    x[5].value = "节假日"
                                    self.hour = 0

        self.wb.save('计算结果.xlsx')

    # 获得加班小时数

    def getHour(self):
        sum4 = 0
        for x in self.ws.rows:
            for y in self.cash.keys():
                if x[0].value == self.name:
                    if x[1].value.strftime("%Y%m%d") == y:
                        sum4 = sum4 + x[7].value
        return sum4

    def sumNum(self, **kw):
        sum3 = 0
        for x in kw:
            for y in self.dict:
                if x == y:
                    sum3 += self.dict[y]
        return sum3

    def dictSeprate(self, target, **kw):
        tt = range(1, len(kw) + 1)
        tup = []
        min1 = 36.5
        c = 0
        minlist = []
        for t in tt:
            tup = list(combinations(tt, t)) + list(tup)

        for x in tup:
            newtup = []
            sum1 = 0
            for y in x:
                c = 0
                for m in kw:
                    c += 1
                    if c == y:
                        newtup.append(m)

            for n in self.dict:
                for k in newtup:
                    if k == n:
                        if self.dict[n] > 0:
                            sum1 += self.dict[n]

            if sum1 <= target:
                if min1 >= target - sum1:
                    min1 = target - sum1
                    minlist = newtup

        for p in minlist:
            for q in self.dict:
                if p == q:
                    self.cash[q] = '转加班费'
        return min1

    def dictSetcash(self, **kw):
        for x in kw:
            self.cash[x] = '转加班费'

    def setConvert(self):
        # 写入是转加班费self.cash还是转串休self.dicts
        for x in self.ws.rows:
            if x[0].value == self.name:
                for y in self.cash.keys():
                    if x[1].value.strftime("%Y%m%d") == y:
                        x[6].value = "转加班费"

                for y in self.dict.keys():
                    if x[1].value.strftime("%Y%m%d") == y:
                        x[6].value = "转串休"
        self.wb.save('计算结果.xlsx')

    def setContents(self, sum, sum_chuan_xiu):
        rng = self.ws2['C2':'AG2']
        for x in rng:
            for y in x:
                for z in self.dict:
                    if y.value.strftime("%Y%m%d") == z:
                        self.ws2.cell(row=name.row, column=y.column).value = self.dict[z]
        self.ws2.cell(row=name.row, column=34).value = sum
        self.ws2.cell(row=name.row, column=35).value = sum_chuan_xiu
        self.wb.save('计算结果.xlsx')

    def jiSuan(self):
        # 获得URL
        self.changeHour()

        # 把self.holiday改回来
        holiday = {}
        weekday = {}
        workday = {}
        for m in self.dict:
            if m in self.holiday:
                holiday[m] = self.dict[m]
            if m in self.weekday:
                weekday[m] = self.dict[m]
            if m in self.workday:
                workday[m] = self.dict[m]

        self.holiday = holiday
        self.weekday = weekday
        self.workday = workday
        # 进行36.5小时判断
        sholiday = self.sumNum(**self.holiday)
        sweekday = self.sumNum(**self.weekday)
        sworkday = self.sumNum(**self.workday)
        sum = sholiday + sweekday + sworkday

        self.getHour()

        if 36.5 - sholiday > 0:
            self.dictSetcash(**self.holiday)
            if 36.5 - sholiday - sweekday > 0:
                self.dictSetcash(**self.weekday)
                if 36.5 - sholiday - sweekday - sworkday > 0:
                    self.dictSetcash(**self.workday)
                else:
                    # 只对workday进行拆分就行
                    self.dictSeprate(36.5 - sholiday - sweekday, **self.workday)
            else:
                # 对workday和weekday同时操作
                min = self.dictSeprate(36.5 - sholiday, **self.weekday)
                min = self.dictSeprate(min, **self.workday)
        else:
            pass

        print('总数据一览：', self.dict)
        print('加班数合计：', round(sum, 2))
        print('转加班小时：', round(self.getHour(), 2))
        sum_chuan_xiu = round(sum - self.getHour(), 2)
        print('转串休小时：', sum_chuan_xiu)
        # 先清空单元格
        for row in self.ws2.iter_rows(min_row=name.row, max_row=name.row
                , min_col=3, max_col=35):
            for cell in row:
                cell.value = None
        self.setContents(sum, sum_chuan_xiu)
        print("转加班费：", sorted(self.cash.keys()))
        for k in self.cash.keys():
            self.dict.pop(k)
        print('转串休假：', sorted(self.dict.keys()))
        self.setConvert()


class Cmacro():
    def __init__(self) -> None:
        self.path = 'C:\\Users\\Administrator\\Documents\\GitHub\\overtime_python\\原始数据.xlsm'

    def dealData(self):
        pass
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(self.path)
        print('START')
        excel.Application.Run("deleteRow")
        wb.Save()
        wb.SaveAs(r'C:\Users\Administrator\Documents\GitHub\overtime_python\计算结果.xlsx', FileFormat=51,
                  ConflictResolution=2)
        wb.Close()
        print('END')


cw = Cwindow()
cw.createWindow()
# 获得工作日和节假日
result = Crili(2023, cw.month).parseHTML()

wb = load_workbook(filename='原始数据.xlsm')
ws = wb['中干']
for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=2):
    for name in row:
        if name.value != None:
            print(name.value, cw.month, '月')
            ji = Count(name, cw.month, result)
            ji.jiSuan()
        else:
            break
