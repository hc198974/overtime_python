# -*- coding: utf-8 -*-
import itertools
import sys
from openpyxl import load_workbook
from demos import *
import copy


class Count_night(object):
    def __init__(self, ids, month, result, wb):
        # 节假日接口(工作日对应结果为 0, 休息日对应结果为 1, 节假日对应的结果为 2 )
        # server_url = "http://www.easybots.cn/api/holiday.php?d="
        self.server_url = "http://tool.bitefu.net/jiari/?d="
        self.wb = wb
        self.ws1 = self.wb["统计表"]
        self.ws2 = self.wb["记录表"]
        self.ws3 = self.wb["明细表"]
        self.ids = ids
        self.id = ""
        self.month = month
        self.cash = {}
        self.records = {}
        self.dictfee_workday = {}
        self.dictfee_weekday = {}
        self.dictfee_holiday = {}
        self.result = result

    def format_date(self, value):
        """将日期值统一格式化为%Y%m%d字符串"""
        if isinstance(value, datetime.datetime):
            return value.strftime("%Y%m%d")
        elif isinstance(value, str):
            # 尝试解析常见日期字符串格式
            for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"]:
                try:
                    return datetime.datetime.strptime(value, fmt).strftime("%Y%m%d")
                except ValueError:
                    continue
            # 如果所有格式都解析失败，返回原始值
            return value
        else:
            return ""

    def floor_half(self, value):
        return int(value * 2) / 2

    def getFee(self, date_key, date):
        if self.result.get(date) == 1.5:
            if date_key == "night":
                return 1

        if self.result.get(date) == 2:
            if date_key == "night":
                return 1.5
            if date_key == "day":
                return 2

        if self.result.get(date) == 3:
            if date_key == "night":
                return 2
            if date_key == "day":
                return 3

    def calday(self, time1, time2):
        temp12 = datetime.datetime.strptime("12:00:00", "%H:%M:%S")
        temp13 = datetime.datetime.strptime("13:00:00", "%H:%M:%S")
        temp8 = datetime.datetime.strptime("8:00:00", "%H:%M:%S")

        hour = 0
        if time2 > time1:
            if time1 < temp8:
                time1 = temp8
            if time2 < temp8:
                time2 = temp8
            if time1 > temp12 and time1 < temp13:
                time1 = temp12
            if time2 > temp12 and time2 < temp13:
                time2 = temp13

            if time1 <= temp12 and time2 >= temp13:
                hour = round(
                    (time2 - time1).seconds / 3600, 2) - 1.5
            else:
                hour = round(
                    (time2 - time1).seconds / 3600, 2) - 0.5

            if hour < 0:
                hour = 0

            return round(hour, 2)

    def writeRecords(self, dictrecords):
        for row in self.ws2.rows:
            if row[1].value == self.id.value:
                rng = self.ws2["C2":"AG2"]
                for x in rng:
                    for y in x:
                        for z in dictrecords:
                            if y.value.strftime("%Y%m%d") == z[0]:
                                temp = self.ws2.cell(
                                    row=self.id.row, column=y.column).value
                                if temp is None:
                                    self.ws2.cell(
                                        row=self.id.row, column=y.column).value = z[3]
                                else:
                                    self.ws2.cell(
                                        row=self.id.row, column=y.column).value = temp+z[3]
                self.ws2.cell(row=self.id.row, column=34).value = sum(
                    list(dictrecords[i][3] for i in range(len(dictrecords))))
            self.ws2.cell(row=self.id.row, column=35).value = sum(
                list(dictrecords[i][3] for i in range(len(dictrecords)))) - 36 if sum(list(dictrecords[i][3] for i in range(len(dictrecords)))) > 36 else 0
        # 同时写入明细表
        for row in self.ws3.rows:
            if row[2].value == self.id.value:
                row[4].value = sum(list(dictrecords[i][3]
                                   for i in range(len(dictrecords))))
                row[5].value = sum(list(dictrecords[i][3]
                                   for i in range(len(dictrecords)) if dictrecords[i][2] == 1.5))
                row[6].value = sum(list(dictrecords[i][3]
                                   for i in range(len(dictrecords)) if dictrecords[i][2] == 2))
                row[7].value = sum(list(dictrecords[i][3]
                                   for i in range(len(dictrecords)) if dictrecords[i][2] == 3))

    def calRecords(self, dictall):
        res = 0
        dictrecords = []
        sorted_data_big = sorted(
            dictall, key=lambda x: (x[2], x[3]), reverse=True)
        count = list(self.result.values()).count(1.5)
        workhours = count*8
        totalhours = sum(x[3] for x in dictall)
        if totalhours >= workhours:
            res = totalhours - workhours
            for x in sorted_data_big:
                if x[3] <= res:
                    res -= x[3]
                    dictrecords.append(x)
                else:
                    hour = x[3] - res
                    dictrecords.append((x[0], x[1], x[2], round(
                        res, 2), "转统计表加班小时数"+str(round(hour, 2))+"小时"))
                    res = 0
                    break

        self.writeRecords(dictrecords)

    def writeDetails(self, dictfee):
        for row in self.ws3.rows:
            if row[2].value == self.id.value:
                row[9].value = self.floor_half(sum(list(dictfee[i][3]
                                                        for i in range(len(dictfee)))))
                row[10].value = self.floor_half(sum(list(dictfee[i][3]
                                                         for i in range(len(dictfee)) if dictfee[i][2] == 1.5)))
                row[11].value = self.floor_half(sum(list(dictfee[i][3]
                                                         for i in range(len(dictfee)) if dictfee[i][2] == 2)))
                row[12].value = self.floor_half(sum(list(dictfee[i][3]
                                                         for i in range(len(dictfee)) if dictfee[i][2] == 3)))
                base_salary = round(row[3].value/21.75/8, 2)
                row[15].value = round(base_salary*1.5*row[10].value, 2)
                row[16].value = round(base_salary*2*row[11].value, 2)
                row[17].value = round(base_salary*3*row[12].value, 2)
                row[14].value = row[15].value+row[16].value + row[17].value

    def calDetails(self, dictall):
        res = 0
        dictfee = []
        list1 = []
        list2 = []
        list3 = []
        list15 = []

        sorted_data_big = sorted(
            dictall, key=lambda x: (x[2], x[3]), reverse=True)
        for x in sorted_data_big:
            if x[2] == 1.5:
                list15.append(x)
            elif x[2] == 1:
                list1.append(x)
            elif x[2] == 2:
                list2.append(x)
            elif x[2] == 3:
                list3.append(x)

        count = list(self.result.values()).count(1.5)
        workhours = count*8
        totalhours = sum(x[3] for x in dictall)
        if totalhours >= workhours:
            res = totalhours - workhours
            if res > 36:
                res = 36

            outer_break = False
            for m in [list3, list2, list15, list1]:
                if sum(x[3] for x in m) <= res:
                    dictfee.extend(m)
                    res -= self.floor_half(sum(x[3] for x in m))
                else:
                    for x in m:
                        if x[3] < res:
                            dictfee.append(x)
                            res -= x[3]
                        else:
                            hour = x[3] - res
                            dictfee.append(
                                (x[0], x[1], x[2], round(res, 2), "转串休"+str(round(hour, 2))+"小时"))
                            res = 0
                            outer_break = True
                            break
                    if outer_break:
                        break

        for x in self.ws1.rows:
            if x[5].value == self.month:
                if x[1].value == self.id.value:
                    dt = self.format_date(x[2].value)
                    for fee in dictfee:
                        if fee[0] == dt:
                            if x[7].value is None:
                                x[7].value = "转加班费"+str(fee[3])
                            else:
                                x[7].value += "，转加班费"+str(fee[3])

                            if len(fee) == 5:
                                x[9].value = fee[4]

                            if fee[2] == 1.5:
                                x[6].value = "值班工作日" if x[6].value is None else x[6].value + "、值班工作日"
                            elif fee[2] == 2:
                                x[6].value = "值班公休日" if x[6].value is None else x[6].value + "、值班公休日"
                            elif fee[2] == 3:
                                x[6].value = "值班节假日" if x[6].value is None else x[6].value + "、值班节假日"
        self.writeDetails(dictfee)

    def writeStatisticsSheet(self):
        temp12 = datetime.datetime.strptime("12:00:00", "%H:%M:%S")
        temp13 = datetime.datetime.strptime("13:00:00", "%H:%M:%S")
        temp8 = datetime.datetime.strptime("8:00:00", "%H:%M:%S")
        temp0 = datetime.datetime.strptime("0:00:00", "%H:%M:%S")
        # 第二天00:00:00，用于计算到当天结束的时间
        temp24 = temp0 + datetime.timedelta(days=1)
        temp1700 = datetime.datetime.strptime("17:00:00", "%H:%M:%S")
        for id in self.ids:
            for x in self.ws1.rows:
                if x[5].value == self.month:
                    if x[1].value == id.value:
                        # 清空没有用的数据
                        x[9].value = None
                        x[8].value = None
                        x[7].value = None
                        x[6].value = None
                        if x[3].value is not None and x[4].value is not None:
                            if x[3].value >= x[4].value:
                                x[4].value = None

        for id in self.ids:
            self.id = id
            dictall = []
            for x in self.ws1.rows:
                if x[5].value == self.month:
                    if x[1].value == id.value:
                        dt = self.format_date(x[2].value)
                        time1 = datetime.datetime.strptime(
                            x[3].value, "%H:%M:%S") if x[3].value is not None else None
                        time2 = datetime.datetime.strptime(
                            x[4].value, "%H:%M:%S") if x[4].value is not None else None

                        # 计算加班小时数
                        # 夜班上半夜
                        if time1 is not None and time2 is None and time1 >= temp12:
                            if time1 < temp1700:
                                time1 = temp1700
                            x[8].value = round(
                                (temp24 - time1).total_seconds() / 3600, 2)
                            fee = self.getFee("night", dt)
                            # 把串休减掉
                            dictall.append(
                                (dt, "night", fee, x[8].value+x[10].value if x[10].value is not None else x[8].value))
                        # 夜班下半夜
                        elif time1 is None and time2 is not None and time2 <= temp12:
                            if time2 > temp8:
                                time2 = temp8
                            x[8].value = round(
                                (time2 - temp0).total_seconds() / 3600, 2)
                            fee = self.getFee("night", dt)
                            dictall.append(
                                (dt, "night", fee, x[8].value+x[10].value if x[10].value is not None else x[8].value))
                        # 一个白天加一个前半夜
                        elif time1 is not None and time1 <= temp12 and time2 is None:
                            day_value = self.calday(time1, temp1700)
                            night_value = round(
                                (temp24 - temp1700).total_seconds() / 3600, 2)
                            x[8].value = day_value + night_value

                            fee = self.getFee("night", dt)
                            dictall.append(
                                (dt, "night", fee, night_value+x[10].value if x[10].value is not None else night_value))
                            fee = self.getFee("day", dt)
                            dictall.append((dt, "day", fee, day_value))
                            x[9].value = "白天加班" + \
                                str(day_value)+"小时，夜班加班"+str(night_value)+"小时"
                        # 一个后半夜+一个白天
                        elif time1 is None and time2 is not None and time2 >= temp13:
                            night_value = round(
                                (temp8 - temp0).total_seconds() / 3600, 2)
                            if time2 > temp1700:
                                time2 = temp1700
                            day_value = self.calday(time1, time2)
                            x[8].value = day_value + night_value
                            fee = self.getFee("night", dt)
                            dictall.append(
                                (dt, "night", fee, night_value+x[10].value if x[10].value is not None else night_value))
                            fee = self.getFee("day", dt)
                            dictall.append((dt, "day", fee, day_value))
                            x[9].value = "白天加班" + \
                                str(day_value)+"小时，夜班加班"+str(night_value)+"小时"
                        # 纯白班
                        elif time1 is not None and time2 is not None:
                            hour = self.calday(time1, temp1700)
                            x[8].value = hour
                            fee = self.getFee("day", dt)
                            dictall.append(
                                (dt, "day", fee, hour+x[10].value if x[10].value is not None else hour))

            # 计入统计表的加班小时数
            self.calRecords(dictall)
            # 计入明细表的加班小时数
            self.calDetails(dictall)

    def jiSuan(self):
        # 将计算结果写入统计表中
        self.writeStatisticsSheet()
