# -*- coding: utf-8 -*-
import itertools
import sys
from openpyxl import load_workbook
from demos import *
import copy


class Count_criterion(object):
    def __init__(self, ids, month, result, wb):
        # 节假日接口(工作日对应结果为 0, 休息日对应结果为 1, 节假日对应的结果为 2 )
        # server_url = "http://www.easybots.cn/api/holiday.php?d="
        self.server_url = "http://tool.bitefu.net/jiari/?d="
        self.wb = wb
        self.ws1 = self.wb["统计表"]
        self.ws2 = self.wb["记录表"]
        self.ids = ids
        self.id = ""
        self.month = month
        self.dictall = {}
        self.weekday = {}
        self.workday = {}
        self.holiday = {}
        self.hour = 0
        self.cash = {}
        self.rest = {}
        self.dictfee = {}
        self.dictfee_workday = {}
        self.dictfee_weekday = {}
        self.dictfee_holiday = {}
        self.result = result

    # `_round_hours_dict` 方法已移除；调用处改为内联四舍五入处理

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

    # 计算统计表里的加班小时数
    def writeStatisticsSheet(self):
        temp17 = datetime.datetime.strptime("17:30:00", "%H:%M:%S")
        temp18 = datetime.datetime.strptime("18:00:00", "%H:%M:%S")
        temp12 = datetime.datetime.strptime("12:00:00", "%H:%M:%S")
        temp13 = datetime.datetime.strptime("13:00:00", "%H:%M:%S")
        temp8 = datetime.datetime.strptime("8:00:00", "%H:%M:%S")
        temp0 = datetime.datetime.strptime("0:00:00", "%H:%M:%S")
        # 第二天00:00:00，用于计算到当天结束的时间
        temp24 = temp0 + datetime.timedelta(days=1)
        temp1700 = datetime.datetime.strptime("17:00:00", "%H:%M:%S")
        self.getUrl()

        for id in self.ids:
            criterion_dict = {}
            sorted_criterion_dict = {}
            temp_dict = {}
            for x in self.ws1.rows:
                if x[5].value == self.month:
                    if x[1].value == id.value:
                        temp = self.format_date(x[2].value)
                        time1 = x[3].value
                        time2 = x[4].value
                        criterion_dict[temp] = (time1, time2)

            # 查找标准班制人员值夜班的情况
            sorted_criterion_dict = dict(sorted(criterion_dict.items()))
            for key, value in sorted_criterion_dict.items():
                # 第一种情况，正常上下班，区分平日和周末
                if value[0] is not None and value[1] is not None:
                    start = datetime.datetime.strptime(
                        value[0], "%H:%M:%S")
                    end = datetime.datetime.strptime(value[1], "%H:%M:%S")
                    hour = 0
                    if key in self.workday.keys():
                        if end >= temp18:
                            hour = round((end - temp17).seconds / 3600, 2)
                            sorted_criterion_dict[key] = value + \
                                (hour, '工作日')
                            temp_dict[key] = hour
                            self.dictall[id.value] = {
                                k: round(v, 2) for k, v in temp_dict.items()}

                    if key in self.weekday.keys() or key in self.holiday.keys():
                        if end > start:
                            if start < temp8:
                                start = temp8
                            if end < temp8:
                                end = temp8
                            if start > temp12 and start < temp13:
                                start = temp12
                            if end > temp12 and end < temp13:
                                end = temp13

                            if start <= temp12 and end >= temp13:
                                hour = round(
                                    (end - start).seconds / 3600, 2) - 1.5
                            else:
                                hour = round(
                                    (end - start).seconds / 3600, 2)-0.5

                            if hour < 0:
                                hour = 0

                            if self.result.get(key) == 2:
                                sorted_criterion_dict[key] = value + \
                                    (hour, '公休日')
                                temp_dict[key] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}
                            elif self.result.get(key) == 3:
                                sorted_criterion_dict[key] = value + \
                                    (hour, '节假日')
                                temp_dict[key] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}

            for n in range(len(sorted_criterion_dict)):
                key1 = list(sorted_criterion_dict.keys())[n]
                value0 = sorted_criterion_dict[key1][0]
                value1 = sorted_criterion_dict[key1][1]
                if n+1 < len(sorted_criterion_dict):
                    key2 = list(sorted_criterion_dict.keys())[n+1]
                    value3 = sorted_criterion_dict[key2][0]
                    value4 = sorted_criterion_dict[key2][1]

                    if value0 is not None and value1 is None:
                        if value3 is None and value4 is not None:
                            start = datetime.datetime.strptime(
                                value0, "%H:%M:%S")
                            end = datetime.datetime.strptime(
                                value4, "%H:%M:%S")
                            hour = 0
                            if start < temp1700:
                                start = temp1700
                            if end > temp8:
                                end = temp8

                            if key1 in self.workday.keys():
                                hour = round(
                                    (temp24 - start).seconds / 3600, 2)
                                sorted_criterion_dict[key1] = value + \
                                    (hour, '工作日(夜班)')
                                temp_dict[key1] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}
                            elif key1 in self.weekday.keys():
                                hour = round(
                                    (temp24 - start).seconds / 3600, 2)
                                sorted_criterion_dict[key1] = value + \
                                    (hour, '公休日(夜班)')
                                temp_dict[key1] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}
                            elif key1 in self.holiday.keys():
                                hour = round(
                                    (temp24 - start).seconds / 3600, 2)
                                sorted_criterion_dict[key1] = value + \
                                    (hour, '节假日(夜班)')
                                temp_dict[key1] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}

                            if key2 in self.workday.keys():
                                hour = round(
                                    (end - temp0).seconds / 3600, 2)
                                sorted_criterion_dict[key2] = value + \
                                    (hour, '工作日(夜班)')
                                temp_dict[key2] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}
                            elif key2 in self.weekday.keys():
                                hour = round(
                                    (end - temp0).seconds / 3600, 2)
                                sorted_criterion_dict[key2] = value + \
                                    (hour, '公休日(夜班)')
                                temp_dict[key2] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}
                            elif key2 in self.holiday.keys():
                                hour = round(
                                    (end - temp0).seconds / 3600, 2)
                                sorted_criterion_dict[key2] = value + \
                                    (hour, '节假日(夜班)')
                                temp_dict[key2] = hour
                                self.dictall[id.value] = {
                                    k: round(v, 2) for k, v in temp_dict.items()}

            for m in self.ws1.rows:
                if m[5].value == self.month:
                    if m[1].value == id.value:
                        s = self.format_date(m[2].value)
                        if s in sorted_criterion_dict.keys():
                            if len(sorted_criterion_dict[s]) == 4:
                                m[8].value = sorted_criterion_dict[s][2]
                                m[6].value = sorted_criterion_dict[s][3]

    def writeRecordsSheet(self):
        # 清空记录表
        for row in self.ws2.iter_rows(min_row=3, max_row=self.ws2.max_row, min_col=3, max_col=35):
            for cell in row:
                cell.value = None

        # 在这里减掉串休使用的小时数
        self.dictfee = copy.deepcopy(self.dictall)
        for row in self.ws1.rows:
            if row[5].value == self.month and row[10].value is not None:
                s = self.format_date(row[2].value)
                if s in self.dictfee.get(row[1].value, {}):
                    if row[10].value > 0:
                        row[10].value = -1*row[10].value
                    self.dictfee[row[1].value][s] += row[10].value

        for id in self.ids:
            # 数据量不大，使用多进程开销大，多线程容易出现错误
            self.id = id

            # 按日期类型分成工作日、公休日、节假日三个字典，并按小时数降序排列
            day_dict = self.dictfee.get(self.id.value, {})
            holiday_items = {
                date_key: hours
                for date_key, hours in day_dict.items()
                if self.result.get(date_key) == 3
            }
            weekday_items = {
                date_key: hours
                for date_key, hours in day_dict.items()
                if self.result.get(date_key) == 2
            }
            workday_items = {
                date_key: hours
                for date_key, hours in day_dict.items()
                if self.result.get(date_key, 1.5) == 1.5
            }

            self.dictfee_holiday[self.id.value] = dict(
                sorted(holiday_items.items(),
                       key=lambda item: item[1], reverse=True)
            )
            self.dictfee_weekday[self.id.value] = dict(
                sorted(weekday_items.items(),
                       key=lambda item: item[1], reverse=True)
            )
            self.dictfee_workday[self.id.value] = dict(
                sorted(workday_items.items(),
                       key=lambda item: item[1], reverse=True)
            )

            # 先扣除加班费小时数较多的日期，保证扣除的串休小时数最少
            def assign_all(category_dict):
                nonlocal remainer, cash_dict
                for date_key, hours in category_dict.items():
                    cash_dict[date_key] = hours
                remainer -= self.floor_half(sum(category_dict.values()))

            # 在扣除加班费小时数较多的日期里逐个扣除串休小时数，直到扣除满36小时或者没有更多的加班小时数可以扣除
            def fill_until_zero(category_dict):
                nonlocal remainer, cash_dict, rest_dict
                for date_key, hours in category_dict.items():
                    if remainer <= 0:
                        rest_dict[date_key] = hours
                        continue
                    if hours <= remainer:
                        cash_dict[date_key] = hours
                        remainer -= hours
                    else:
                        cash_dict[date_key] = remainer
                        rest_dict[date_key] = hours - remainer
                        remainer = 0
                return remainer == 0

            remainer = 36
            cash_dict = {}
            rest_dict = {}
            sorted_dict = dict(
                sorted(day_dict.items(), key=lambda item: item[0]))

            if sum(self.dictfee_holiday[self.id.value].values()) >= remainer:
                fill_until_zero(self.dictfee_holiday[self.id.value])
            else:
                assign_all(self.dictfee_holiday[self.id.value])
                if remainer > 0:
                    if sum(self.dictfee_weekday[self.id.value].values()) >= remainer:
                        fill_until_zero(
                            self.dictfee_weekday[self.id.value])
                    else:
                        assign_all(self.dictfee_weekday[self.id.value])
                        if remainer > 0:
                            fill_until_zero(
                                self.dictfee_workday[self.id.value])

            for y in self.ws1.rows:
                if y[8].value is not None:
                    if y[1].value == self.id.value and y[8].value > 0:
                        s = self.format_date(y[2].value)
                        if s in cash_dict:
                            y[7].value = "转加班费"
                        if s in cash_dict and s in rest_dict:
                            y[7].value = "转加班费" + str(round(cash_dict.get(s, 0), 2)) + \
                                "小时、" + "转串休" + \
                                str(round(rest_dict.get(s, 0), 2)) + "小时"
                        if s not in cash_dict:
                            y[7].value = "转串休"

            self.cash[self.id.value] = cash_dict
            self.rest[self.id.value] = rest_dict
            rng = self.ws2["C2":"AG2"]
            for x in rng:
                for y in x:
                    for z in sorted_dict:
                        if y.value.strftime("%Y%m%d") == z:
                            self.ws2.cell(row=self.id.row, column=y.column).value = sorted_dict[
                                z
                            ]
            self.ws2.cell(row=self.id.row, column=34).value = sum(
                list(sorted_dict.values()))
            self.ws2.cell(row=self.id.row, column=35).value = sum(
                list(sorted_dict.values())) - 36 if sum(list(sorted_dict.values())) > 36 else 0

    def writeDetailsSheet(self):
        ws3 = self.wb["明细表"]
        # 写入加班小时数和加班费
        for row in ws3.iter_rows(min_row=3, max_row=ws3.max_row, min_col=3, max_col=3):
            for cell in row:
                # 要区分夜班和非夜班人员，夜班人员的加班费计算方式不同
                if cell.value is not None:
                    # 写入总加班小时
                    overtime_dict = self.dictall.get(cell.value, {})
                    # 合计
                    ws3.cell(row=cell.row, column=5).value = round(
                        sum(overtime_dict.values()), 2)
                    # 工作日加班小时数
                    ws3.cell(
                        row=cell.row,
                        column=6,
                        value=round(sum(
                            hours for date_key, hours in overtime_dict.items()
                            if self.result.get(date_key) == 1.5
                        ), 2),
                    )
                    # 公休日加班小时数
                    ws3.cell(
                        row=cell.row,
                        column=7,
                        value=round(sum(
                            hours for date_key, hours in overtime_dict.items()
                            if self.result.get(date_key) == 2
                        ), 2),
                    )
                    # 节假日加班小时数
                    ws3.cell(
                        row=cell.row,
                        column=8,
                        value=round(sum(
                            hours for date_key, hours in overtime_dict.items()
                            if self.result.get(date_key) == 3
                        ), 2),
                    )
                    # 写入串休扣除数
                    # 从统计表中汇总该 id 在本月的串休扣除（统计表第10列，索引9）并写入明细表第9列
                    # !!!这里需注意，是不是夜班人员，串休扣除数都是这个功能填上的
                    rest_sum = sum(
                        (r[10].value if r[10].value is not None else 0)
                        for r in self.ws1.rows
                        if r[5].value == self.month and r[1].value == cell.value
                    )
                    ws3.cell(row=cell.row, column=9, value=rest_sum)

                    # 写入扣除串休后计入加班费的小时数!!!
                    fee_dict = self.cash.get(cell.value, {})

                    # 扣除串休后计入加班费的小时数
                    workday_fee = self.floor_half(
                        sum(
                            hours
                            for date_key, hours in fee_dict.items()
                            if self.result.get(date_key) == 1.5
                        )
                    )
                    weekday_fee = self.floor_half(
                        sum(
                            hours
                            for date_key, hours in fee_dict.items()
                            if self.result.get(date_key) == 2
                        )
                    )
                    holiday_fee = self.floor_half(
                        sum(
                            hours
                            for date_key, hours in fee_dict.items()
                            if self.result.get(date_key) == 3
                        )
                    )

                    ws3.cell(
                        row=cell.row,
                        column=11,
                        value=workday_fee,
                    )
                    ws3.cell(
                        row=cell.row,
                        column=12,
                        value=weekday_fee,
                    )
                    ws3.cell(
                        row=cell.row,
                        column=13,
                        value=holiday_fee,
                    )

                    # 小计：等于右侧三列之和
                    ws3.cell(
                        row=cell.row,
                        column=10,
                        value=workday_fee + weekday_fee + holiday_fee,
                    )

                    # 加班费基数，成本科要求在这保存两位小数
                    base_salary = round(
                        ws3.cell(row=cell.row, column=4).value/21.75/8, 2)
                    # 工作日加班费
                    ws3.cell(
                        row=cell.row,
                        column=16,
                        value=round(workday_fee * base_salary * 1.5, 2),
                    )

                    # 公休日加班费
                    ws3.cell(
                        row=cell.row,
                        column=17,
                        value=round(weekday_fee * base_salary * 2, 2),
                    )

                    # 节假日加班费
                    ws3.cell(
                        row=cell.row,
                        column=18,
                        value=round(holiday_fee * base_salary * 3, 2),
                    )

                    # 小计（加班费总额）：第16/17/18列之和，保留两位小数，写入第15列
                    total_pay = (
                        (ws3.cell(row=cell.row, column=16).value or 0)
                        + (ws3.cell(row=cell.row, column=17).value or 0)
                        + (ws3.cell(row=cell.row, column=18).value or 0)
                    )
                    ws3.cell(row=cell.row, column=15,
                             value=round(total_pay, 2))

    def jiSuan(self):
        # 将计算结果写入统计表中
        self.writeStatisticsSheet()

        # 将计算结果写入记录表中
        self.writeRecordsSheet()

        # 将计算结果写入明细表中
        self.writeDetailsSheet()
