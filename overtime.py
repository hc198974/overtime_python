# -*- coding: utf-8 -*-
import itertools
import sys
from openpyxl import load_workbook
from demos import *
import copy


class Count(object):
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
        self.result = result
        self.night_num = ['60836', 'Q4642']

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
        def calculateCriterion(self):
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
                special_dict = {}
                sorted_special_dict = {}
                temp_dict = {}
                if not id.value in self.night_num:
                    for x in self.ws1.rows:
                        if x[5].value == self.month:
                            if x[1].value == id.value:
                                temp = self.format_date(x[2].value)
                                time1 = x[3].value
                                time2 = x[4].value
                                special_dict[temp] = (time1, time2)

                # 查找标准班制人员值夜班的情况
                sorted_special_dict = dict(sorted(special_dict.items()))
                for key, value in sorted_special_dict.items():
                    # 第一种情况，正常上下班，区分平日和周末
                    if value[0] is not None and value[1] is not None:
                        start = datetime.datetime.strptime(
                            value[0], "%H:%M:%S")
                        end = datetime.datetime.strptime(value[1], "%H:%M:%S")
                        hour = 0
                        if key in self.workday.keys():
                            if end >= temp18:
                                hour = round((end - temp17).seconds / 3600, 2)
                                sorted_special_dict[key] = value + \
                                    (hour, '工作日')
                                temp_dict[key] = hour

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
                                    sorted_special_dict[key] = value + \
                                        (hour, '公休日')
                                    temp_dict[key] = hour
                                elif self.result.get(key) == 3:
                                    sorted_special_dict[key] = value + \
                                        (hour, '节假日')
                                    temp_dict[key] = hour
                                    self.dictall[id.value] = temp_dict

                for n in range(len(sorted_special_dict)):
                    key1 = list(sorted_special_dict.keys())[n]
                    value0 = sorted_special_dict[key1][0]
                    value1 = sorted_special_dict[key1][1]
                    if n+1 < len(sorted_special_dict):
                        key2 = list(sorted_special_dict.keys())[n+1]
                        value3 = sorted_special_dict[key2][0]
                        value4 = sorted_special_dict[key2][1]

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
                                    sorted_special_dict[key1] = value + \
                                        (hour, '工作日(夜班)')
                                    temp_dict[key1] = hour
                                    self.dictall[id.value] = temp_dict
                                elif key1 in self.weekday.keys():
                                    hour = round(
                                        (temp24 - start).seconds / 3600, 2)
                                    sorted_special_dict[key1] = value + \
                                        (hour, '公休日(夜班)')
                                    temp_dict[key1] = hour
                                    self.dictall[id.value] = temp_dict
                                elif key1 in self.holiday.keys():
                                    hour = round(
                                        (temp24 - start).seconds / 3600, 2)
                                    sorted_special_dict[key1] = value + \
                                        (hour, '节假日(夜班)')
                                    temp_dict[key1] = hour
                                    self.dictall[id.value] = temp_dict

                                if key2 in self.workday.keys():
                                    hour = round(
                                        (end - temp0).seconds / 3600, 2)
                                    sorted_special_dict[key2] = value + \
                                        (hour, '工作日(夜班)')
                                    temp_dict[key2] = hour
                                    self.dictall[id.value] = temp_dict
                                elif key2 in self.weekday.keys():
                                    hour = round(
                                        (end - temp0).seconds / 3600, 2)
                                    sorted_special_dict[key2] = value + \
                                        (hour, '公休日(夜班)')
                                    temp_dict[key2] = hour
                                    self.dictall[id.value] = temp_dict
                                elif key2 in self.holiday.keys():
                                    hour = round(
                                        (end - temp0).seconds / 3600, 2)
                                    sorted_special_dict[key2] = value + \
                                        (hour, '节假日(夜班)')
                                    temp_dict[key2] = hour
                                    self.dictall[id.value] = temp_dict

                for m in self.ws1.rows:
                    if m[5].value == self.month:
                        if m[1].value == id.value:
                            s = self.format_date(m[2].value)
                            if s in sorted_special_dict.keys():
                                if len(sorted_special_dict[s]) == 4:
                                    m[8].value = sorted_special_dict[s][2]
                                    m[6].value = sorted_special_dict[s][3]

        def calculateNight(self):
            temp17 = datetime.datetime.strptime("17:00:00", "%H:%M:%S")
            temp8 = datetime.datetime.strptime("8:00:00", "%H:%M:%S")
            temp0 = datetime.datetime.strptime("0:00:00", "%H:%M:%S")
            # 第二天00:00:00，用于计算到当天结束的时间
            temp24 = temp0 + datetime.timedelta(days=1)

            # 是工作日还是周末在calculateCriterion函数里已经判断过了，这里直接计算夜班加班时间
            for id in self.ids:
                total_dict = {}
                total = 0
                if id.value in self.night_num:
                    for x in self.ws1.rows:
                        if x[5].value == self.month:
                            if x[1].value == id.value and x[1].value in self.night_num:
                                temp = self.format_date(x[2].value)
                                time1 = x[3].value
                                time2 = x[4].value
                                # 第一种情况：上半夜计算方式
                                if time1 != "" and time1 is not None:
                                    if type(time1) == str:
                                        time1 = datetime.datetime.strptime(
                                            time1, "%H:%M:%S")
                                    elif type(time1) == datetime.time:
                                        time1 = datetime.datetime.strptime(
                                            (time1.strftime("%H:%M:%S")), "%H:%M:%S"
                                        )  # 先把datetime.time格式转换为str再转换为datetime.datetime

                                    if time1 > temp17:
                                        total_dict[temp] = round(
                                            (temp24 - time1).seconds / 3600, 2)
                                    else:
                                        total_dict[temp] = round(
                                            (temp24 - temp17).seconds / 3600, 2)

                                # 第二种情况：下半夜计算方式
                                if time1 is None and time2 != "" and time2 is not None:
                                    if type(time2) == str:
                                        time2 = datetime.datetime.strptime(
                                            time2, "%H:%M:%S")
                                    elif type(time2) == datetime.time:
                                        time2 = datetime.datetime.strptime(
                                            (time2.strftime("%H:%M:%S")), "%H:%M:%S"
                                        )  # 先把datetime.time格式转换为str再转换为datetime.datetime

                                    if time2 < temp8:
                                        total_dict[temp] = round(
                                            (time2 - temp0).seconds / 3600, 2)
                                    else:
                                        total_dict[temp] = round(
                                            (temp8 - temp0).seconds / 3600, 2)

                    total = len(self.workday)*8
                    night_total = sum(list(total_dict.values()))
                    if night_total > total:
                        sorted_dict = dict(
                            sorted(total_dict.items(), key=lambda item: (
                                self.result.get(item[0], 1), item[0]))
                        )
                        truncated_dict = {}
                        list_total = list(sorted_dict.items())
                        for m in range(len(sorted_dict)):
                            total -= list_total[m][1]
                            if total < 0:
                                truncated_dict[list_total[m][0]] = total*(-1)
                                truncated_dict.update(
                                    dict(list_total[m+1:]))
                                break

                        # 存入总加班时间
                        self.dictall[id.value] = truncated_dict
                        # 将加班时间写入统计表
                        for x in self.ws1.rows:
                            if x[5].value == self.month:
                                if x[1].value == id.value and x[1].value in self.night_num:
                                    s = self.format_date(x[2].value)
                                    if s in truncated_dict.keys():
                                        x[8].value = truncated_dict[s]
                                        if s in self.workday.keys():
                                            x[6].value = "工作日"
                                        elif s in self.weekday.keys():
                                            x[6].value = "公休日"
                                        elif s in self.holiday.keys():
                                            x[6].value = "节假日"
                                    else:
                                        x[8].value = 0
                                        if s in self.workday.keys():
                                            x[6].value = "工作日"
                                        elif s in self.weekday.keys():
                                            x[6].value = "公休日"
                                        elif s in self.holiday.keys():
                                            x[6].value = "节假日"
        # 标准班制加班时间计算
        calculateCriterion(self)
        # 夜班
        calculateNight(self)

    def writeRecordsSheet(self):
        # 清空记录表
        for row in self.ws2.iter_rows(min_row=3, max_row=self.ws2.max_row, min_col=3, max_col=35):
            for cell in row:
                cell.value = None

        # 在这里减掉串休使用的小时数
        self.dictfee = copy.deepcopy(self.dictall)
        for row in self.ws1.rows:
            if row[5].value == self.month and row[9].value is not None:
                s = self.format_date(row[2].value)
                if s in self.dictfee.get(row[1].value, {}):
                    if row[9].value > 0:
                        row[9].value = -1*row[9].value
                    self.dictfee[row[1].value][s] += row[9].value

        for id in self.ids:
            # 数据量不大，使用多进程开销大，多线程容易出现错误
            self.id = id
            category_order = {3: 0, 2: 1, 1.5: 2}
            sorted_dict = dict(
                sorted(
                    self.dictfee.get(self.id.value, {}).items(),
                    key=lambda item: (
                        category_order.get(self.result.get(item[0], 1.5), 3),
                        -item[1],
                    ),
                )
            )
            remainer = 36
            cash_dict = {}
            rest_dict = {}
            if sum(list(sorted_dict.values())) > remainer:
                for x in sorted_dict:
                    if remainer > 0 and remainer - sorted_dict[x] > 0:
                        cash_dict.update(
                            {x: sorted_dict[x]})
                        remainer -= sorted_dict[x]
                    elif remainer > 0 and remainer - sorted_dict[x] <= 0:
                        cash_dict.update({x: remainer})
                        rest_dict.update(
                            {x: sorted_dict[x] - remainer})
                        remainer -= sorted_dict[x]
                    elif remainer <= 0:
                        rest_dict.update({x: sorted_dict[x]})
                        remainer -= sorted_dict[x]

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
                            if not s in cash_dict and s in rest_dict:
                                y[7].value = "转串休"

                self.cash[self.id.value] = cash_dict
                self.rest[self.id.value] = rest_dict
            else:
                self.cash[self.id.value] = sorted_dict
                for m in self.ws1.rows:
                    if m[5].value == self.month:
                        if m[8].value is not None:
                            if m[1].value == self.id.value and m[8].value > 0:
                                m[7].value = "转加班费"

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

                    # 写入扣除串休后计入加班费的小时数
                    fee_dict = self.cash.get(cell.value, {})
                    # 小计
                    ws3.cell(
                        row=cell.row,
                        column=9,
                        value=round(sum(
                            hours for date_key, hours in fee_dict.items()
                        ), 2),
                    )
                    # 工作日扣除串休后计入加班费的小时数
                    ws3.cell(
                        row=cell.row,
                        column=10,
                        value=round(
                            sum(
                                hours
                                for date_key, hours in fee_dict.items()
                                if self.result.get(date_key) == 1.5
                            ),
                            2,
                        ),
                    )
                    # 公休日扣除串休后计入加班费的小时数
                    ws3.cell(
                        row=cell.row,
                        column=11,
                        value=round(
                            sum(
                                hours
                                for date_key, hours in fee_dict.items()
                                if self.result.get(date_key) == 2
                            ),
                            2,
                        ),
                    )
                    # 节假日扣除串休后计入加班费的小时数
                    ws3.cell(
                        row=cell.row,
                        column=12,
                        value=round(
                            sum(
                                hours
                                for date_key, hours in fee_dict.items()
                                if self.result.get(date_key) == 3
                            ),
                            2,
                        ),
                    )

                    # 加班费基数，成本科要求在这保存两位小数
                    base_salary = round(
                        ws3.cell(row=cell.row, column=4).value/21.75/8, 2)
                    # 工作日加班费
                    if cell.value not in self.night_num:
                        ws3.cell(
                            row=cell.row,
                            column=15,
                            value=round(
                                sum(
                                    hours
                                    for date_key, hours in fee_dict.items()
                                    if self.result.get(date_key) == 1.5
                                )
                                * base_salary * 1.5,
                                2,
                            ),
                        )
                    else:
                        ws3.cell(
                            row=cell.row,
                            column=15,
                            value=round(
                                sum(
                                    hours
                                    for date_key, hours in fee_dict.items()
                                    if self.result.get(date_key) == 1.5
                                )
                                * base_salary * 1,  # 夜班工作日按1倍计算
                                2,
                            ),
                        )

                    # 公休日加班费
                    if cell.value not in self.night_num:
                        ws3.cell(
                            row=cell.row,
                            column=16,
                            value=round(
                                sum(
                                    hours
                                    for date_key, hours in fee_dict.items()
                                    if self.result.get(date_key) == 2
                                )
                                * base_salary * 2,
                                2,
                            ),
                        )
                    else:
                        ws3.cell(
                            row=cell.row,
                            column=16,
                            value=round(
                                sum(
                                    hours
                                    for date_key, hours in fee_dict.items()
                                    if self.result.get(date_key) == 2
                                )
                                * base_salary * 1,  # 夜班公休日按1倍计算
                                2,
                            ),
                        )

                    # 节假日加班费
                    if cell.value not in self.night_num:
                        ws3.cell(
                            row=cell.row,
                            column=17,
                            value=round(
                                sum(
                                    hours
                                    for date_key, hours in fee_dict.items()
                                    if self.result.get(date_key) == 3
                                )
                                * base_salary * 3,
                                2,
                            ),
                        )
                    else:
                        ws3.cell(
                            row=cell.row,
                            column=17,
                            value=round(
                                sum(
                                    hours
                                    for date_key, hours in fee_dict.items()
                                    if self.result.get(date_key) == 3
                                )
                                * base_salary * 2,  # 夜班节假日按2倍计算
                                2,
                            ),
                        )

    def jiSuan(self):
        # 将计算结果写入统计表中
        self.writeStatisticsSheet()

        # 将计算结果写入记录表中
        self.writeRecordsSheet()

        # 将计算结果写入明细表中
        self.writeDetailsSheet()


if __name__ == "__main__":
    cw = Cwindow()
    if not cw.createWindow():
        print("已取消计算：窗口已关闭")
        sys.exit(0)

    # 获得工作日和节假日
    result = Crili(2026, cw.month).parseHTML()
    start = time.perf_counter()
    wb = load_workbook(filename="计算结果.xlsx")
    ws = wb["记录表"]
    # 用id取代name，识别职号
    ids = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=2):
        for id in row:
            if id.value is not None:
                ids.append(id)
            else:
                break
    ji = Count(ids, cw.month, result, wb)
    ji.jiSuan()
    wb.save("计算结果.xlsx")

    print("运行时间：", time.perf_counter() - start)
