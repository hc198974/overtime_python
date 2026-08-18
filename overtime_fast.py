# -*- coding: utf-8 -*-
"""
overtime_fast.py —— overtime.py 的高速重写版

与 overtime.py 功能完全一致（输入/输出同一份「计算结果.xlsx」的「汇总表」/「中干」），
但用两种方法替换原实现的性能瓶颈：
  1. 时长计算：原实现是「每人 × 全表」双重循环 + 逐格 openpyxl 读写（O(人数×行数)）；
     本版改为 pandas/numpy 向量化（O(行数)），常量只解析一次。
  2. 36 小时封顶分配：原实现用 itertools.combinations 枚举子集（最坏 O(2^n)）；
     本版改为 0-1 背包动态规划（O(n×3600/人)），且因日类权重 3>2>1.5 严格有序，
     其最优解在数学上等价于原「分组贪心 + 枚举求最大」逻辑，结果一致。

年份推断：跟随 GUI 所选月份 + 当前年份（datetime.now().year）。
"""
import datetime
import time

import numpy as np
import pandas as pd

from demos import Cwindow, Crili

# 常量时间只解析一次（避免在原实现的 hot loop 内重复解析）
_T17 = datetime.datetime.strptime("17:30:00", "%H:%M:%S")
_T18 = datetime.datetime.strptime("18:00:00", "%H:%M:%S")
_T12 = datetime.datetime.strptime("12:00:00", "%H:%M:%S")
_T13 = datetime.datetime.strptime("13:00:00", "%H:%M:%S")
_T08 = datetime.datetime.strptime("8:00:00", "%H:%M:%S")
_TODAY = datetime.date(2000, 1, 1)

# 秒级常量
_S17 = 17 * 3600 + 30 * 60
_S18 = 18 * 3600
_S12 = 12 * 3600
_S13 = 13 * 3600
_S08 = 8 * 3600


def _fmt_date(v):
    """将日期值统一格式化为 %Y%m%d 字符串（与日历字典 key 对齐）。"""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y%m%d")
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return s


def calc_duration(df, calendar):
    """复刻 overtime.py changeHour 的时长公式，全程 numpy 向量化（O(行数)）。

    返回带「时长」「节假日」与内部列「_dt」(日历类型 1.5/2/3) 的 DataFrame。
    """
    signin = pd.to_datetime(df["签到"].astype(str), format="%H:%M:%S", errors="coerce")
    signout = pd.to_datetime(df["签出"].astype(str), format="%H:%M:%S", errors="coerce")
    sb = signin.dt.hour * 3600 + signin.dt.minute * 60 + signin.dt.second
    xb = signout.dt.hour * 3600 + signout.dt.minute * 60 + signout.dt.second
    d = df["日报日期"].map(lambda v: calendar.get(_fmt_date(v), 0)).to_numpy()

    dur = np.zeros(len(df), dtype=float)
    label = np.full(len(df), "", dtype=object)

    valid = sb.notna().to_numpy() & xb.notna().to_numpy() & (xb.to_numpy() > sb.to_numpy()) & (d > 0)
    sb = sb.to_numpy().astype(float)
    xb = xb.to_numpy().astype(float)

    # 工作日：仅 签出 > 18:00 计，时长 = 签出 - 17:30
    wd = valid & (d == 1.5)
    dur_wd = np.where(xb > _S18, np.clip(xb - _S17, 0, None), 0.0)
    dur = np.where(wd, dur_wd, dur)
    label = np.where(wd, "工作日", label)

    # 周末 / 节假日
    we = valid & (d != 1.5)
    sb_n = np.where(sb > _S08, sb, _S08)                 # 签入 <=8:00 → 8:00
    sb_n = np.where((sb_n > _S12) & (sb_n < _S13), _S12, sb_n)  # 12<签入<13 → 12:00
    xb_n = np.where((xb > _S12) & (xb < _S13), _S13, xb)  # 12<签出<13 → 13:00
    delta = xb_n - sb_n
    deduct = np.where(
        xb_n <= _S12, 1800.0,
        np.where(xb_n >= _S13,
                 np.where(sb_n <= _S12, 5400.0, 1800.0),
                 0.0))
    h = delta - deduct
    mid = (xb_n > _S12) & (xb_n < _S13)                  # 12:00~13:00 之间不计
    dur_we = np.where(mid, 0.0, np.where(h >= 0, h, 0.0))
    dur = np.where(we, dur_we, dur)
    label = np.where(we, "节假日", label)

    out = df.copy()
    out["时长"] = np.round(dur / 3600, 2)
    out["节假日"] = label
    out["_dt"] = d
    return out


def allocate_36h(sub):
    """对单人的加班明细做 36h 封顶分配，返回应「转加班费」的行索引集合。

    0-1 背包 DP：容量 3600(=36h×100)，权重 = 时长×100，价值 = 时长×日类权重×100。
    日类权重 3>2>1.5 严格有序，故 DP 最优解等价于原「分组贪心 + 枚举求最大」逻辑。
    """
    items = sub[(sub["时长"] > 0) & (sub["_dt"] > 0)].copy()
    if len(items) == 0:
        return set()
    if items["时长"].sum() <= 36:
        return set(items.index.tolist())
    hours = items["时长"].to_numpy(dtype=float)
    dtypes = items["_dt"].to_numpy(dtype=float)
    weights = (hours * 100).round().astype("int64")
    values = (hours * dtypes * 100).round().astype("int64")
    cap = 3600
    n = len(weights)
    dp = [0] * (cap + 1)
    for i in range(n):
        w = int(weights[i])
        v = int(values[i])
        for c in range(cap, w - 1, -1):
            if dp[c - w] + v > dp[c]:
                dp[c] = dp[c - w] + v
    c = cap
    selected = set()
    for i in range(n - 1, -1, -1):
        w = int(weights[i])
        v = int(values[i])
        if c >= w and dp[c] == dp[c - w] + v:
            selected.add(items.index[i])
            c -= w
    return selected


def generate_summary_table(df):
    """重建「中干」汇总表：序号 / 姓名 / 各日期 / 合计 / 可串休时间。"""
    pivot = df.pivot_table(index="姓名", columns="日报日期",
                           values="时长", aggfunc="sum", fill_value=0)
    pivot["合计"] = pivot.sum(axis=1)
    ksx = (df[df["加班或串休"] == "转串休"]
           .groupby("姓名")["时长"].sum()
           .rename("可串休时间"))
    summary = pivot.reset_index()
    summary = summary.merge(ksx, on="姓名", how="left").fillna({"可串休时间": 0})
    summary.insert(0, "序号", range(1, len(summary) + 1))
    return summary


def run_core(df, calendar, month):
    """核心计算：仅对所选月份的行重算时长与 36h 分配，其余月份行保持原样。"""
    # 兼容「刚清洗过的空表」：节假日/加班或串休 可能被推断为 float，先转 object
    for c in ("节假日", "加班或串休"):
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(object)
    mask = df["月份"] == month
    work = df[mask].copy()
    work = calc_duration(work, calendar)

    work["加班或串休"] = ""
    for _, g in work.groupby("姓名"):
        sel = allocate_36h(g)
        mask_sel = work.index.isin(sel)
        mask_pos = work["时长"] > 0
        work.loc[mask_sel & mask_pos, "加班或串休"] = "转加班费"
        work.loc[~mask_sel & mask_pos, "加班或串休"] = "转串休"

    # 把本月重算结果写回原 DataFrame（不影响其它月份行）
    df.loc[mask, "时长"] = work["时长"].to_numpy()
    df.loc[mask, "节假日"] = work["节假日"].to_numpy()
    df.loc[mask, "加班或串休"] = work["加班或串休"].to_numpy()

    summary = generate_summary_table(work)
    return df, summary


def main(month, year):
    """主流程：抓日历 -> 读表 -> 计算 -> 写回（与 overtime.py 同文件同表名）。"""
    calendar = Crili(year, month).parseHTML()
    df = pd.read_excel("计算结果.xlsx", sheet_name="汇总表")
    df, summary = run_core(df, calendar, month)
    with pd.ExcelWriter("计算结果.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="汇总表")
        summary.to_excel(writer, index=False, sheet_name="中干")
    return df, summary


if __name__ == "__main__":
    cw = Cwindow()
    cw.createWindow()                 # GUI 选择月份（沿用 overtime.py 交互）
    year = datetime.datetime.now().year  # 年份跟随当前年份自动推断
    start = time.perf_counter()
    main(cw.month, year)
    print("运行时间：", time.perf_counter() - start)
