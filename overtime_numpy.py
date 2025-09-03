# -*- coding: utf-8 -*-
import datetime
import time
import pandas as pd
import numpy as np
from demos import *


def overtime_cal(datas, result):
    """加班计算"""
    modified_chunk = []
    for i in range(len(datas)//len(result)):
        chunk = datas[i*len(result):(i*len(result))+(len(result)-1)]
        chunk = dealdata(chunk, result)
        chunk = maximize_value(chunk)
        modified_chunk.append(chunk)

    flat_list = [item for sublist in modified_chunk for item in sublist]
    # 转换为DataFrame
    df = pd.DataFrame(flat_list, columns=[
        '姓名', '日报日期', '上班时间', '下班时间',
        '月份', '节假日', '加班或串休', '时长'
    ])
    return df

# 1-0背包问题，使用动态规划解决


def maximize_value(result, max_sum=36000):
    if sum([row[7] for row in result]) <= max_sum / 1000:
        for row in result:
            if row[7] > 0:
                row[6] = '转加班'
        return result

    # 使用滚动数组优化
    weights = [int(row[7] * 1000) for row in result]
    values = [int(row[7] * row[5] * 1000) for row in result]
    n = len(weights)

    # 创建一维 DP 数组
    dp = [0] * (max_sum + 1)

    # 填充 DP 数组
    for i in range(n):
        for w in range(max_sum, weights[i] - 1, -1):
            dp[w] = max(dp[w], dp[w - weights[i]] + values[i])

    # 回溯找到选择的物品
    w = max_sum
    selected_items = set()
    for i in range(n - 1, -1, -1):
        if w >= weights[i] and dp[w] == dp[w - weights[i]] + values[i]:
            selected_items.add(i)
            w -= weights[i]

    # 更新结果
    for i, row in enumerate(result):
        if i in selected_items and row[7] > 0:
            row[6] = '转加班'
        elif row[7] > 0:
            row[6] = '转串休'

    return result


def dealdata(chunk, calendar):
    def calculate_time(sb, xb, isweekend):
        temp17 = datetime.datetime.strptime("17:30:00", "%H:%M:%S").time()
        temp18 = datetime.datetime.strptime("18:00:00", "%H:%M:%S").time()
        temp12 = datetime.datetime.strptime("12:00:00", "%H:%M:%S").time()
        temp13 = datetime.datetime.strptime("13:00:00", "%H:%M:%S").time()
        temp8 = datetime.datetime.strptime("8:00:00", "%H:%M:%S").time()

        if pd.isna(sb) or pd.isna(xb):
            return 0
        if isinstance(sb, str):
            sb = datetime.datetime.strptime(sb, "%H:%M:%S").time()
        if isinstance(xb, str):
            xb = datetime.datetime.strptime(xb, "%H:%M:%S").time()
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
                    return max(delta-1.5, 0)  # 防止出现负数
                else:
                    return max(delta-0.5, 0)
            return 0

    # 先导入节假日
    chunk[:, 5] = np.vectorize(calendar.get, otypes=[np.float64])(
        chunk[:, 1]).astype(float)
    # 第二步计算加班小时数
    chunk[:, 7] = np.vectorize(calculate_time, otypes=[np.float64])(
        chunk[:, 2], chunk[:, 3], chunk[:, 5])
    chunk = chunk.tolist()

    return chunk


def generate_summary_table(df):
    # 使用 groupby 替代 pivot_table，性能更好
    summary_df = df.groupby(['姓名', '日报日期'])['时长'].sum().unstack(fill_value=0)
    summary_df['合计'] = summary_df.sum(axis=1)
    return summary_df.reset_index()


if __name__ == "__main__":
    cw = Cwindow()
    cw.createWindow()
    with requests.Session() as session:
        # 获得工作日和节假日
        calendar = Crili(2025, cw.month).parseHTML()
        start = time.perf_counter()
        df = pd.read_excel('计算结果.xlsx', parse_dates=['日报日期'])

        df['日报日期'] = df['日报日期'].dt.strftime('%Y%m%d')
        datas = df.to_numpy()
        sorted_indices = np.argsort(datas[:, 0])  # 获取排序索引
        datas = datas[sorted_indices]  # 按索引重新排列数组
        df = overtime_cal(datas, calendar)
        summary_df = generate_summary_table(df)
        # 多表导出到excel
        with pd.ExcelWriter("计算结果.xlsx") as writer:
            df.to_excel(writer, index=False,
                        sheet_name='明细', engine='openpyxl')
            summary_df.to_excel(writer, index=False,
                                sheet_name='汇总', engine='openpyxl')
    print("运行时间：", time.perf_counter() - start)
