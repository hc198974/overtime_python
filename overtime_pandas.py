# -*- coding: utf-8 -*-
import itertools
import pandas as pd
from demos import *


def custom_gettime(df):
    temp17 = datetime.datetime.strptime("17:30:00", "%H:%M:%S").time()
    temp18 = datetime.datetime.strptime("18:00:00", "%H:%M:%S").time()
    temp12 = datetime.datetime.strptime("12:00:00", "%H:%M:%S").time()
    temp13 = datetime.datetime.strptime("13:00:00", "%H:%M:%S").time()
    temp8 = datetime.datetime.strptime("8:00:00", "%H:%M:%S").time()

    def calculate_time(row):
        sb = row['上班时间']
        xb = row['下班时间']
        if pd.isna(sb) or pd.isna(xb):
            return 0
        if isinstance(sb, str):
            sb = datetime.datetime.strptime(sb, "%H:%M:%S").time()
        if isinstance(xb, str):
            xb = datetime.datetime.strptime(xb, "%H:%M:%S").time()

        if row['节假日'] == 1.5:
            if xb >= sb and sb <= temp17:
                if xb >= temp18:
                    return round((datetime.datetime.combine(datetime.date.today(), xb) - datetime.datetime.combine(
                        datetime.date.today(), temp17)).seconds / 3600, 2)
            return 0
        else:
            delta = round((datetime.datetime.combine(datetime.date.today(
            ), xb) - datetime.datetime.combine(datetime.date.today(), sb)).seconds / 3600, 2)
            if delta > 0.5:
                if sb < temp8:
                    sb = temp8
                if sb > temp12 and sb < temp13:
                    sb = temp13
                if xb > temp12 and xb < temp13:
                    xb = temp13
                # 计算加班时间
                delta = round((datetime.datetime.combine(datetime.date.today(
                ), xb) - datetime.datetime.combine(datetime.date.today(), sb)).seconds / 3600, 2)
                if xb >= temp13 and sb <= temp12:
                    return delta - 1.5
                else:
                    return delta - 0.5
            return 0

    df['时长'] = df.apply(calculate_time, axis=1)
    return df


def getgroup(dict, group3, group2, group1):
    def get_start(array, remainer):
        i = 1
        s = 0
        while i <= len(array):
            if sum(array[:i][::1]) >= remainer:
                s = i - 1
                break
            else:
                i += 1
        return s

    def get_end(array, remainer):
        i = 1
        s = 0
        while i <= len(array):
            if sum(array[:i][::-1]) >= remainer:
                s = i
                break
            else:
                i += 1
        return s

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
            array = sorted(
                group2.loc[group2['时长'] > 0, '时长'].tolist(), reverse=True)
            array_start = get_start(array, remainer)
            array_end = get_end(array, remainer)
            for i in range(array_start, array_end):
                combinations = list(
                    itertools.combinations(list(group2[group2['时长'] != 0].index), i))
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
            array = sorted(
                group1.loc[group1['时长'] > 0, '时长'].tolist(), reverse=True)
            array_start = get_start(array, remainer)
            array_end = get_end(array, remainer)
            for i in range(array_start, array_end):
                combinations = list(
                    itertools.combinations(list(group1.index), i))
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


def main(calendars):
    start = time.perf_counter()
    list = []
    # 读取Excel文件，默认第一个表《汇总表》
    df = pd.read_excel('计算结果.xlsx', parse_dates=['日报日期'])
    df['日报日期'] = df['日报日期'].dt.strftime('%Y%m%d')
    df.drop('节假日', axis=1, inplace=True)
    calendars_df = pd.DataFrame(calendars.items(), columns=['日报日期', '节假日'])
    df = df.merge(calendars_df, on='日报日期', how='left')
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

    with pd.ExcelWriter("计算结果.xlsx") as writer:
        df.to_excel(writer, index=False, sheet_name='明细', engine='openpyxl')
        summary_df.to_excel(writer, index=False,
                            sheet_name='汇总', engine='openpyxl')
    print("时间：", time.perf_counter() - start)


if __name__ == "__main__":
    cw = Cwindow()
    cw.createWindow()
    calendars = Crili(2025, cw.month).parseHTML()
    start = time.perf_counter()
    main(calendars)
    print("运行时间：", time.perf_counter() - start)
