from openpyxl import load_workbook
from demos import Crili
import datetime
import calendar
import logging
from copy import deepcopy

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 时间常量（秒）
T0, T8, T12, T13, T17, T18, T24 = (
    0, 8 * 3600, 12 * 3600, 13 * 3600, 17 * 3600, 18 * 3600, 24 * 3600,
)

# 夜班类型映射表（基础映射）
NIGHT_TYPE_MAP = {
    "unknown": "未知",
    "normal": "正常",
    "back": "后半夜",
    "front": "前半夜",
    "front+back": "前后半夜",
    "day_front": "白天+前半夜",
    "day_back": "后半夜+白天",
    "day": "白班",
}


def get_night_type_display(night_type: str, wd) -> str:
    """根据夜班类型和工作日类型返回显示文字。
    - 混合模式 day+front / back+day 仅在休息日/节假日（wd=2/3）显示"白天"部分
    - 工作日（wd=1.5）下，17:00 前不算加班，混合模式只保留前后夜段
    Args:
        night_type: 班次类型字符串
        wd: 工作日类型（1.5=工作日, 2=休息日, 3=节假日）
    Returns:
        统计表 E 列显示文字
    """
    if night_type == "day+front":
        if wd in (2, 3):
            return "白天+前半夜"
        return "前半夜"
    if night_type == "back+day":
        if wd in (2, 3):
            return "后半夜+白天"
        return "后半夜"
    return NIGHT_TYPE_MAP.get(str(night_type), str(night_type))


def to_sec(t: str) -> int:
    """时间字符串转秒数
    Args:
        t: 时间字符串，格式 HH:MM:SS
    Returns:
        对应的秒数
    """
    h, m, s = map(int, t.split(":"))
    return h * 3600 + m * 60 + s


def load_workbook_safe(filename: str):
    """安全加载Excel文件，添加异常处理
    Args:
        filename: Excel文件名
    Returns:
        工作簿对象
    Raises:
        FileNotFoundError: 文件不存在
        PermissionError: 文件被占用
    """
    try:
        return load_workbook(filename=filename)
    except FileNotFoundError:
        logging.error(f"文件 {filename} 不存在")
        raise
    except PermissionError:
        logging.error(f"文件 {filename} 被占用，请先关闭后重试")
        raise


crili = Crili(2026, datetime.datetime.now().month - 1)
weekday = crili.parseHTML()


# 工号 → 姓名 映射（由 calculate_main 读取进出场记录时填充）
dict_id_to_name = {}


def calculate_dict_overtime(dict_in_out: dict, emp_id: str, d: str) -> list:
    """计算标准人员日加班时长
    Args:
        dict_in_out: 进出场记录字典
        emp_id: 工号
        d: 日期字符串
    Returns:
        [加班时长, 夜班类型]
    """
    

    records = dict_in_out[emp_id][d]
    if not records:
        return [0, "unknown"]
    if d not in weekday:
        return [0, "unknown"]
    wd = weekday[d]

    def overlap(a: int, b: int, x: int, y: int) -> int:
        """计算两个区间的重叠秒数"""
        return max(0, min(b, y) - max(a, x))

    in_count = sum(1 for _, direction in records if direction == "in")
    out_count = sum(1 for _, direction in records if direction == "out")

    # 班次类型说明：
    # - 前半夜加班 = 17:00-24:00 加班
    # - 后半夜加班 = 0:00-8:00 加班
    # - 白天 = 8:00-17:00（白班 8 点起步，进厂 < 8:00 按 8:00 算）
    # - 跨月份边界：若 d 是该月第一天则前一日视为无数据（可默认未出厂）；
    #              若 d 是该月最后一天则后一日视为无数据（可默认正常出厂）。
    #
    # 情况一：进厂记录多于出厂记录 → 白天+前半夜（混合模式）
    # 判定当天是否有白天+前半夜加班：看第二天该人员最后一条记录
    # 第二天最后记录为出厂 → 有白天+前半夜加班（次日 0:00 之后出厂）
    # 第二天最后记录为进厂 → 无白天+前半夜加班，记录不闭合
    # 第二天无数据：
    #   - d 是该月最后一天 → 默认正常出厂，视为有白天+前半夜加班
    #   - 否则 → 视为有进无出的连续不闭合记录，无法判定
    if in_count > out_count:
        next_d = (
            datetime.datetime.strptime(d, "%Y%m%d")
            + datetime.timedelta(days=1)
        ).strftime("%Y%m%d")
        next_records = dict_in_out.get(emp_id, {}).get(next_d, [])
        if next_records and next_records[-1][1] == "in":
            name = dict_id_to_name.get(emp_id, "")
            logging.warning(
                f"{name}（{emp_id}）{d} 进厂多于出厂，且第二天{next_d}最后记录为进厂，无法判定白天+前半夜加班"
            )
            night_type = "unknown"
            return [0, night_type]
        if not next_records:
            d_obj = datetime.datetime.strptime(d, "%Y%m%d")
            last_day = calendar.monthrange(d_obj.year, d_obj.month)[1]
            if d_obj.day != last_day:
                name = dict_id_to_name.get(emp_id, "")
                logging.warning(
                    f"{name}（{emp_id}）{d} 进厂多于出厂，且第二天{next_d}无数据（非月末最后一天），无法判定白天+前半夜加班"
                )
                night_type = "unknown"
                return [0, night_type]
        # 补齐 24:00 作为白天+前半夜加班的最晚结束时间
        records = list(records) + [("24:00:00", "out")]
        # 情况一：仅 1 条 in + 补齐 24:00，必然有白天+前半夜
        night_type = "day+front"
    # 情况二：进厂记录少于出厂记录 → 后半夜+白天（混合模式）
    # 判定当天是否有后半夜+白天加班：看前一天该人员最后一条记录
    # 前一天最后记录为进厂（即前一天未出厂，延续到今天）→ 有后半夜+白天加班
    # 前一天最后记录为出厂 → 无后半夜+白天加班依据，记录不闭合
    # 前一天无数据：
    #   - d 是该月第一天 → 默认未出厂，视为有后半夜+白天加班
    #   - 否则 → 视为有出无进的连续不闭合记录，无法判定
    elif in_count < out_count:
        prev_d = (
            datetime.datetime.strptime(d, "%Y%m%d") - datetime.timedelta(days=1)
        ).strftime("%Y%m%d")
        prev_records = dict_in_out.get(emp_id, {}).get(prev_d, [])
        if prev_records and prev_records[-1][1] == "out":
            name = dict_id_to_name.get(emp_id, "")
            logging.warning(
                f"{name}（{emp_id}）{d} 出厂多于进厂，且前一天{prev_d}最后记录为出厂，无法判定后半夜+白天加班"
            )
            night_type = "unknown"
            return [0, night_type]
        if not prev_records:
            d_obj = datetime.datetime.strptime(d, "%Y%m%d")
            if d_obj.day != 1:
                name = dict_id_to_name.get(emp_id, "")
                logging.warning(
                    f"{name}（{emp_id}）{d} 出厂多于进厂，且前一天{prev_d}无数据（非月初第一天），无法判定后半夜+白天加班"
                )
                night_type = "unknown"
                return [0, night_type]
        # 补齐 0:00 作为后半夜+白天加班的延续开始时间（前一天最后一刻进厂）
        records = [("00:00:00", "in")] + list(records)
        # 根据最后一条 out 时间判定班次类型
        out_times = [to_sec(t) for t, direction in records if direction == "out"]
        if out_times and max(out_times) >= T8:
            # out ≥ 8:00 → 真正有后半夜+白天
            night_type = "back+day"
        else:
            # out < 8:00 → 仅后半夜加班，无白天
            night_type = "back"
    # 情况三：进出相等 → 分两种子情况处理
    # 3a) 最大出厂 < 最小入厂 → 后半夜 + 前半夜
    #     判定后半夜加班：看前一天该人员最后一条记录
    #         前一天最后记录为进厂（未出厂）→ 有后半夜加班
    #         前一天最后记录为出厂 → 无后半夜加班，记录不闭合
    #         前一天无数据 → 默认未出厂，视为有后半夜加班
    #     判定前半夜加班：看第二天该人员最后一条记录
    #         第二天最后记录为出厂 → 有前半夜加班
    #         第二天最后记录为进厂 → 无前半夜加班，记录不闭合
    #         第二天无数据 → 默认正常出厂，视为有前半夜加班
    # 3b) 首条 in 且末条 out → 标准班
    # 3c) 其他不闭合情况 → unknown
    elif in_count == out_count:
        out_times = [to_sec(t) for t, direction in records if direction == "out"]
        in_times = [to_sec(t) for t, direction in records if direction == "in"]
        if max(out_times) < min(in_times):
            # 后半夜验证：前一天最后记录
            prev_d = (
                datetime.datetime.strptime(d, "%Y%m%d") - datetime.timedelta(days=1)
            ).strftime("%Y%m%d")
            prev_records = dict_in_out.get(emp_id, {}).get(prev_d, [])
            if prev_records and prev_records[-1][1] == "out":
                logging.warning(
                    f"{d} 最大出厂早于最小入厂，但前一天{prev_d}最后记录为出厂，无法判定后半夜加班"
                )
                night_type = "unknown"
                return [0, night_type]
            # 前半夜验证：第二天最后记录
            next_d = (
                datetime.datetime.strptime(d, "%Y%m%d")
                + datetime.timedelta(days=1)
            ).strftime("%Y%m%d")
            next_records = dict_in_out.get(emp_id, {}).get(next_d, [])
            if next_records and next_records[-1][1] == "in":
                logging.warning(
                    f"{d} 最大出厂早于最小入厂，但第二天{next_d}最后记录为进厂，无法判定前半夜加班"
                )
                night_type = "unknown"
                return [0, night_type]
            night_type = "front+back"
            if wd == 1.5:
                intervals = [(T0, T8), (T17, T24)]
            elif wd in (2, 3):
                intervals = [(T0, T12), (T13, T24)]
            else:
                intervals = []
            dict_overtime_sec = 0
            for lo, hi in intervals:
                dict_overtime_sec += overlap(0, min(out_times), lo, hi)
                dict_overtime_sec += overlap(max(in_times), T24, lo, hi)
            return [round(max(0, dict_overtime_sec / 3600), 2), night_type]
        elif records[0][1] == "in" and records[-1][1] == "out":
            night_type = "normal"
        else:
            logging.warning(f"{d} 进出场记录不闭合，且不匹配已知情况")
            night_type = "unknown"
            return [0, night_type]
    else:
        logging.warning(f"{d} 进出场记录不闭合，且不匹配已知情况")
        night_type = "unknown"
        return [0, night_type]

    adjust = night_type == "normal"

    if wd == 1.5:
        if night_type == "back":
            # 后半夜加班：0:00-8:00
            intervals = [(T0, T8)]
        elif night_type == "front":
            # 前半夜加班：17:00-24:00
            intervals = [(T17, T24)]
        elif night_type == "day+front":
            # 白天+前半夜（工作日）：仅前半夜 17:00-24:00（17:00 前不算加班）
            intervals = [(T17, T24)]
        elif night_type == "back+day":
            # 后半夜+白天（工作日）：仅后半夜 0:00-8:00（17:00 前不算加班）
            intervals = [(T0, T8)]
        else:
            intervals = [(T18, T24)]
    elif wd in (2, 3):
        if night_type == "back":
            intervals = [(T0, T8)]
        elif night_type == "front":
            intervals = [(T17, T24)]
        elif night_type == "day+front":
            # 白天+前半夜：8:00-17:00（休息日/节假日）+ 17:00-24:00
            intervals = [(T8, T12), (T13, T17), (T17, T24)]
        elif night_type == "back+day":
            # 后半夜+白天（休息日/节假日）：0:00-8:00 + 8:00-17:00
            intervals = [(T0, T8), (T8, T12), (T13, T17)]
        else:
            intervals = [(T8, T12), (T13, T24)]
    else:
        intervals = []

    dict_overtime_sec = 0
    in_time = None
    # 白班 8 点起步：day+front 模式下 in < 8:00 视为 8:00
    round_in_to_t8 = night_type == "day+front"
    for t, direction in records:
        ts = to_sec(t)
        if direction == "in":
            in_time = ts
            if round_in_to_t8 and in_time < T8:
                in_time = T8
        elif direction == "out" and in_time is not None:
            out_time = ts
            for lo, hi in intervals:
                dict_overtime_sec += overlap(in_time, out_time, lo, hi)
            in_time = None

     # 支持特殊工号的自定义算法（曲书成在这两天是白班模式）
    if emp_id == "Q6007":
        if d in ["20260801", "20260808"]:
            try:
                dict_overtime_sec=0
                intervals = [(T0, T12), (T13, T24)]
                for t, direction in records:                    
                    ts = to_sec(t)
                    if direction == "in":
                        in_time = ts
                    elif direction == "out" and in_time is not None:
                        out_time = ts
                        
                        for lo, hi in intervals:
                            dict_overtime_sec += overlap(in_time, out_time, lo, hi)
                        in_time = None
                        
            except Exception as e:
                logging.warning(f"special overtime standard failed for {emp_id} {d}: {e}, fallback to default")
    
    if adjust and dict_overtime_sec > 0:
        if wd == 1.5:
            dict_overtime_sec += 1800
        elif wd in (2, 3):
            dict_overtime_sec -= 1800

     
   
    return [round(max(0, dict_overtime_sec / 3600), 2), night_type]


def calculate_dict_overtime_night(dict_in_out: dict, emp_id: str, d: str) -> list:
    """夜班加班计算逻辑，融合自 overtime_night.py
    Args:
        dict_in_out: 进出场记录字典
        emp_id: 工号
        d: 日期字符串
    Returns:
        [日期, 总加班时长元组, 系数元组, 夜班类型]
    """

    records = dict_in_out[emp_id][d]
    
    if not records:
        return [d, (0, 0), (0, 0), "unknown"]
    if d not in weekday:
        return [d, (0, 0), (0, 0), "unknown"]
    wd = weekday[d]

    def calday(time1: int, time2: int) -> float:
        """计算白天加班（扣除午休）"""
        hour = 0
        if time2 > time1:
            if time1 < T8:
                time1 = T8
            if time2 < T8:
                time2 = T8
            if T12 < time1 < T13:
                time1 = T12
            if T12 < time2 < T13:
                time2 = T13
            if time1 <= T12 and time2 >= T13:
                hour = round((time2 - time1) / 3600, 2) - 1.5
            else:
                hour = round((time2 - time1) / 3600, 2) - 0.5
            if hour < 0:
                hour = 0
            return round(hour, 2)
        return 0

    def get_coe(wd_val: float) -> tuple:
        """根据班次获取加班和时间"""
        if wd_val == 1.5:
            return (1.5, 1.5)
        elif wd_val == 2:
            return (2, 1.5)
        elif wd_val == 3:
            return (3, 2)
        else:
            return (0.0, 0.0)

    in_times = [to_sec(t) for t, direction in records if direction == "in"]
    out_times = [to_sec(t) for t, direction in records if direction == "out"]

    time1 = in_times[0] if in_times else None
    time2 = out_times[-1] if out_times else None

    total_hours = (0.0, 0.0)
    coe = (0.0, 0.0)
    night_type = "unknown"

    # 场景 1: 夜班上半夜 (有进无出，且进厂时间 >= 12:00)
    if time1 is not None and time2 is None and time1 >= T12:
        night_type = "front"
        if time1 < T17:
            time1 = T17
        total_hours = (0.0, round((T24 - time1) / 3600, 2))
        coe = get_coe(wd)

    # 场景 2: 夜班下半夜 (无进有出，且出厂时间 <= 12:00)
    elif time1 is None and time2 is not None and time2 <= T12:
        night_type = "back"
        if time2 > T8:
            time2 = T8
        total_hours = (0.0, round((time2 - T0) / 3600, 2))
        coe = get_coe(wd)

    # 场景 3: 白天 + 前半夜 (有进无出，且进厂时间 <= 12:00)
    elif time1 is not None and time1 <= T12 and time2 is None:
        night_type = "day_front"
        day_value = calday(time1, T17)
        night_value = round((T24 - T17) / 3600, 2)
        total_hours = (day_value, night_value)
        coe = get_coe(wd)

    # 场景 4: 后半夜 + 白天 (无进有出，且出厂时间 >= 13:00)
    elif time1 is None and time2 is not None and time2 >= T13:
        night_type = "day_back"
        night_value = round((T8 - T0) / 3600, 2)
        if time2 > T17:
            time2 = T17
        day_value = calday(T8, time2)
        total_hours = (day_value, night_value)
        coe = get_coe(wd)

    # 场景 5: 纯白班 (有进有出)
    elif time1 is not None and time2 is not None and time1<time2:
        night_type = "day"
        hour = calday(time1, min(time2, T17))
        total_hours = (hour, 0.0)
        coe = get_coe(wd)     
    
    # 场景 6: 连续上夜班，前半夜+后半夜
    elif time1 is not None and time2 is not None and time2<time1:
        night_type = "front+back"
        if time1 < T17:
            time1 = T17
        if time2 > T8:
            time2 = T8
        total_hours = (0.0, round((T24 - time1) / 3600 + (time2 - T0) / 3600, 2))
        coe = get_coe(wd)
    else:
        logging.warning(f"{d} 进出场记录异常，无法判定夜班类型")
        return [d, (0.0, 0.0), (0.0, 0.0), "unknown"]
   
    
    
    return [d, total_hours, coe, night_type]


def format_date(value) -> str:
    """将日期值统一格式化为%Y%m%d字符串
    Args:
        value: 日期值，支持datetime.datetime、str等类型
    Returns:
        格式化后的日期字符串
    """
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y%m%d")
    elif isinstance(value, str):
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"]:
            try:
                return datetime.datetime.strptime(value, fmt).strftime("%Y%m%d")
            except ValueError:
                continue
        return value
    else:
        return ""


def dedup_consecutive(records: list) -> list:
    """去重连续同向记录
    连续 in 保留最小（最早），连续 out 保留最大（最晚）
    Args:
        records: 已按时间升序排列的记录列表
    Returns:
        去重后的记录列表
    """
    if not records:
        return records
    result = [records[0]]
    for t, direction in records[1:]:
        last_t, last_dir = result[-1]
        if direction == last_dir:
            if direction == "in":
                continue
            else:
                result[-1] = (t, direction)
        else:
            result.append((t, direction))
    return result


def floor_half(hours: float) -> float:
    """向下取 0.5 的整数倍
    Args:
        hours: 小时数
    Returns:
        向下取整后的小时数
    """
    return int(hours * 2) / 2


def apply_row10_adjustments(
    rows,
    dict_used_overtime=None,
    dict_overtime=None,
    dict_overtime_night=None,
    night_ids=None,
):
    """读取统计表 row[10] 的非 0 调整值，并把它们应用到加班结果。"""
    adjustments = {}
    for row in rows:
        id_value = row[1].value
        date_value = format_date(row[2].value)
        value = row[10].value
        if not id_value:
            continue
        if value in (None, 0, 0.0, "", False):
            continue

        adjustments.setdefault(id_value, {})[date_value] = value

        if dict_used_overtime is None or dict_overtime is None:
            continue

        if id_value in dict_used_overtime and date_value in dict_used_overtime[id_value]:
            dict_used_overtime[id_value][date_value][0] = (
                dict_overtime[id_value][date_value][0] + value
            )

        if night_ids is not None and id_value in night_ids:
            for i, entry in enumerate(dict_overtime_night[id_value]):
                if entry[0] != date_value:
                    continue
                dayvalue, nightvalue = entry[1]
                if nightvalue + value > 0:
                    nightvalue += value
                else:
                    s = nightvalue + value
                    nightvalue = 0
                    dayvalue += s
                dict_overtime_night[id_value][i] = [
                    entry[0],
                    (dayvalue, nightvalue),
                    entry[2],
                    entry[3],
                ]
    return adjustments


def collect_adjustments(*args, **kwargs):
    """兼容旧调用入口，实际委托给 apply_row10_adjustments。"""
    return apply_row10_adjustments(*args, **kwargs)


def get_total_adjustment_hours(total_chuanxiu, id_value):
    """汇总某个工号在总调整值字典中的和。"""
    if not total_chuanxiu or id_value not in total_chuanxiu:
        return 0.0
    return round(sum(total_chuanxiu[id_value].values()), 2)


def _set_if_nonzero(cell, value) -> None:
    """值非 0 时才写入单元格；0 / 0.0 / None / '' 跳过
    Args:
        cell: Excel单元格对象
        value: 要写入的值
    """
    if value in (0, 0.0, None, ""):
        return
    cell.value = value


def _write_mingxi_row(row: list, emp_id: str, total: float, work: float,
                     rest: float, holiday: float, pay_work: float,
                     pay_rest: float, pay_holiday: float,
                     adjustment_total: float) -> None:
    """写入明细表单行的公共逻辑
    Args:
        row: 工作表行对象
        emp_id: 工号
        total: 总加班时长
        work: 工作日加班
        rest: 休息日加班
        holiday: 节假日加班
        pay_work: 工作日转加班费
        pay_rest: 休息日转加班费
        pay_holiday: 节假日转加班费
        adjustment_total: 串休调整总和
    """
    _set_if_nonzero(row[4], round(total, 2))
    _set_if_nonzero(row[5], round(work, 2))
    _set_if_nonzero(row[6], round(rest, 2))
    _set_if_nonzero(row[7], round(holiday, 2))
    _set_if_nonzero(row[8], adjustment_total)
    floored_pay_work = floor_half(round(pay_work, 2))
    floored_pay_rest = floor_half(round(pay_rest, 2))
    floored_pay_holiday = floor_half(round(pay_holiday, 2))
    _set_if_nonzero(row[10], floored_pay_work)
    _set_if_nonzero(row[11], floored_pay_rest)
    _set_if_nonzero(row[12], floored_pay_holiday)
    _set_if_nonzero(
        row[9],
        round(floored_pay_work + floored_pay_rest + floored_pay_holiday, 2),
    )

    # 加班费金额：工资基数 = round(基本工资 / 21.75 / 8, 2)
    # P列工作日1.5倍 / Q列公休日2倍 / R列节假日3倍，O列为三者之和，加班费取整数
    basic_salary = row[3].value
    if isinstance(basic_salary, (int, float)) and basic_salary > 0:
        base_rate = round(basic_salary / 21.75 / 8, 2)
        p_value = round(base_rate * 1.5 * floored_pay_work)
        q_value = round(base_rate * 2 * floored_pay_rest)
        r_value = round(base_rate * 3 * floored_pay_holiday)
        o_value = round(p_value + q_value + r_value)
        _set_if_nonzero(row[15], p_value)
        _set_if_nonzero(row[16], q_value)
        _set_if_nonzero(row[17], r_value)
        _set_if_nonzero(row[14], o_value)


def write_dict_overtime_to_excel(
    ids,
    night_ids,
    dict_in_out,
    dict_overtime,
    dict_overtime_night,
    dict_overtime_night_original,
    dict_used_overtime,
    dict_used_overtime_night,
    total_chuanxiu=None,
):
    """将加班结果写入Excel文件（统计表、记录表、明细表）"""
    wb = load_workbook_safe("计算结果.xlsx")

    # ============ 统计表 ============
    ws = wb["统计表"]
    all_ids_set = set(ids) | set(night_ids)

    # 1) 清空目标 id 之前写入的单元格
    clear_cols_stat = {3, 4, 6, 7, 8, 9}
    for row in ws.iter_rows(min_row=2):
        if row[1].value not in all_ids_set:
            continue
        for col in clear_cols_stat:
            row[col].value = None

    # 2) 为夜班人员构建按日期索引的加班数据（调整后，用于加班费/串休）
    night_overtime_by_date = {}
    for night_id in night_ids:
        night_overtime_by_date[night_id] = {}
        if night_id in dict_overtime_night:
            for entry in dict_overtime_night[night_id]:
                date = entry[0]
                total_hours = sum(entry[1])
                night_type = entry[3]
                night_overtime_by_date[night_id][date] = [total_hours, night_type]

    # 2.1) 为夜班人员构建按日期索引的原始加班数据（未调整，用于统计表"时长"列）
    night_overtime_by_date_original = {}
    for night_id in night_ids:
        night_overtime_by_date_original[night_id] = {}
        if night_id in dict_overtime_night_original:
            for entry in dict_overtime_night_original[night_id]:
                date = entry[0]
                total_hours = sum(entry[1])
                night_type = entry[3]
                night_overtime_by_date_original[night_id][date] = [total_hours, night_type]
            logging.info(f"夜班人员 {night_id}: {len(dict_overtime_night[night_id])} 条加班记录")

    # 3) 为夜班人员构建按日期索引的加班费/串休数据（累加同一天的多条记录）
    # dict_used_overtime_night 结构: {id: [[date, t, coe, night_type, x, y], ...]}
    # x = 转加班费, y = 转串休
    night_pay_by_date = {}
    for night_id in night_ids:
        night_pay_by_date[night_id] = {}
        if night_id in dict_used_overtime_night:
            for entry in dict_used_overtime_night[night_id]:
                date = entry[0]
                x = entry[4]
                y = entry[5]
                if date in night_pay_by_date[night_id]:
                    night_pay_by_date[night_id][date][0] += x
                    night_pay_by_date[night_id][date][1] += y
                else:
                    night_pay_by_date[night_id][date] = [x, y]

    # 4) 重新写入所有人员（标准 + 夜班）
    for row in ws.rows:
        id = row[1].value
        if id not in all_ids_set:
            continue

        dt = format_date(row[2].value)
        if dt not in dict_in_out.get(id, {}):
            continue

        row[3].value = str(dict_in_out[id][dt])

        is_night = id in night_ids

        if is_night:
            # 加班费/串休使用调整后的值
            if dt in night_overtime_by_date.get(id, {}):
                _, night_type_for_pay = night_overtime_by_date[id][dt]
            else:
                night_type_for_pay = "unknown"
            # 时长列使用原始值（不含adjustments）
            if dt in night_overtime_by_date_original.get(id, {}):
                total_hours, night_type = night_overtime_by_date_original[id][dt]
            else:
                total_hours, night_type = 0, "unknown"
        else:
            if dt in dict_overtime.get(id, {}):
                total_hours, night_type = dict_overtime[id][dt]
            else:
                total_hours, night_type = 0, "unknown"

        wd_value = weekday.get(dt)
        row[4].value = get_night_type_display(str(night_type), wd_value)

        if dt in weekday:
            if weekday[dt] == 1.5:
                row[6].value = "工作日"
            elif weekday[dt] == 2:
                row[6].value = "休息日"
            elif weekday[dt] == 3:
                row[6].value = "节假日"

        _set_if_nonzero(row[8], total_hours)
        row[9].value = "夜班" if is_night else "标准"

        if is_night:
            if dt in night_pay_by_date.get(id, {}):
                pay, comp_off = night_pay_by_date[id][dt]
                if pay > 0 or comp_off > 0:
                    row[7].value = f"转加班费{round(pay, 2)}；转串休{round(comp_off, 2)}"
        else:
            if dt in dict_used_overtime.get(id, {}):
                pay = dict_used_overtime[id][dt][2]
                comp_off = dict_used_overtime[id][dt][3]
                if pay > 0 or comp_off > 0:
                    row[7].value = f"转加班费{round(pay, 2)}；转串休{round(comp_off, 2)}"

    # ============ 记录表 ============
    ws = wb["记录表"]
    date_col = {}
    for cell in ws[2]:
        if isinstance(cell.value, datetime.datetime):
            date_col[cell.value.strftime("%Y%m%d")] = cell.column
    ah_col = next((c.column for c in ws[2] if c.value == "合计"), None)
    ai_col = next((c.column for c in ws[2] if c.value == "可串休时间"), None)
    write_cols_jilu = set(date_col.values())
    if ah_col:
        write_cols_jilu.add(ah_col)
    if ai_col:
        write_cols_jilu.add(ai_col)

    # 1) 清空目标 id 之前写入的单元格
    all_ids_jilu = set(ids) | set(night_ids)
    for row in ws.iter_rows(min_row=3):
        if row[1].value not in all_ids_jilu:
            continue
        for col in write_cols_jilu:
            row[col - 1].value = None

    # 2) 写入标准人员加班记录
    for row in ws.iter_rows(min_row=3):
        id = row[1].value
        if id not in ids or id not in dict_used_overtime:
            continue
        total_hours = 0.0
        total_comp_off = 0.0
        for d, col in date_col.items():
            if d in dict_used_overtime[id]:
                hours, _, pay, comp_off = dict_used_overtime[id][d]
                _set_if_nonzero(row[col - 1], hours)
                total_hours += hours
                total_comp_off += comp_off
        if ah_col:
            _set_if_nonzero(row[ah_col - 1], round(total_hours, 2))
        if ai_col:
            _set_if_nonzero(row[ai_col - 1], round(total_comp_off, 2))

    # 3) 写入夜班人员的加班记录
    for row in ws.iter_rows(min_row=3):
        id = row[1].value
        if id not in night_ids or id not in dict_used_overtime_night:
            continue
        night_agg = {}
        for entry in dict_used_overtime_night[id]:
            date, t, coe, night_type, x, y = entry
            if date not in night_agg:
                night_agg[date] = {"pay": 0.0, "comp_off": 0.0}
            night_agg[date]["pay"] += x
            night_agg[date]["comp_off"] += y
        night_total_pay_comp = 0.0
        night_total_comp_off = 0.0
        for d, col in date_col.items():
            if d in night_agg:
                day_value = night_agg[d]["pay"] + night_agg[d]["comp_off"]
                _set_if_nonzero(row[col - 1], round(day_value, 2))
                night_total_pay_comp += day_value
                night_total_comp_off += night_agg[d]["comp_off"]
        if ah_col:
            _set_if_nonzero(row[ah_col - 1], round(night_total_pay_comp, 2))
        if ai_col:
            _set_if_nonzero(row[ai_col - 1], round(night_total_comp_off, 2))

    # ============ 明细表 ============
    ws = wb["明细表"]
    write_cols_mingxi = {5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18}

    # 1) 清空目标 id 的写入区域
    all_ids_mingxi = set(ids) | set(night_ids)
    for row in ws.iter_rows(min_row=4):
        if row[2].value not in all_ids_mingxi:
            continue
        for col in write_cols_mingxi:
            row[col - 1].value = None

    # 2) 重新写入（标准人员）
    for row in ws.iter_rows(min_row=4):
        id = row[2].value
        if id not in ids or id not in dict_overtime:
            continue
        total = work = rest = holiday = 0.0

        for d, entry in dict_overtime[id].items():
            hours, _ = entry
            total += hours
            wd = weekday.get(d, 0)
            if wd == 1.5:
                work += hours
            elif wd == 2:
                rest += hours
            elif wd == 3:
                holiday += hours

        adjustment_total = get_total_adjustment_hours(total_chuanxiu, id)

        pay_total = pay_work = pay_rest = pay_holiday = 0.0
        for d in dict_used_overtime[id]:
            pay_total += dict_used_overtime[id][d][2]
            wd = weekday.get(d, 0)
            if wd == 1.5:
                pay_work += dict_used_overtime[id][d][2]
            elif wd == 2:
                pay_rest += dict_used_overtime[id][d][2]
            elif wd == 3:
                pay_holiday += dict_used_overtime[id][d][2]

        _write_mingxi_row(row, id, total, work, rest, holiday,
                         pay_work, pay_rest, pay_holiday, adjustment_total)

    # 3) 写入夜班人员的加班记录
    # 夜班人员E-H列：截断后 x+y 的总和（= 加班费+串休 = 调整前总加班 - work_hours）
    #            J-M列：使用截断后 x 的值（加班费，不超过36小时）
    # 数据流：dict_overtime_night_original → [+workhours] → dict_overtime_night → [36h截断] → dict_used_overtime_night
    # E-H列 = 截断后 x+y 之和（= 调整前总加班 - work_hours）
    # J-M列 = 截断后 x 的值
    for row in ws.iter_rows(min_row=4):
        id = row[2].value
        if id not in night_ids or id not in dict_used_overtime_night:
            continue
        
        # E-H列：使用截断后的 dict_used_overtime_night 中 x+y 之和
        # 结构: [date, t, coe, night_type, x, y]
        night_data = dict_used_overtime_night[id]
        total = work = rest = holiday = 0.0  # E-H列（x+y 按 coe 分类）
        pay_work = pay_rest = pay_holiday = 0.0  # J-M列（仅 x）

        for entry in night_data:
            date, t, coe, night_type, x, y = entry
            hours = x + y  # 加班费 + 串休
            total += hours
            if coe == 1.5:
                work += hours
                pay_work += x
            elif coe == 2:
                rest += hours
                pay_rest += x
            elif coe == 3:
                holiday += hours
                pay_holiday += x
        
        total = round(total, 2)
        work = round(work, 2)
        rest = round(rest, 2)
        holiday = round(holiday, 2)

        adjustment_total = get_total_adjustment_hours(total_chuanxiu, id)

        _write_mingxi_row(row, id, total, work, rest, holiday,
                         pay_work, pay_rest, pay_holiday, adjustment_total)

    wb.save(filename="计算结果.xlsx")
    wb.close()


def calculate_36h_truncation(valid_ids: list, dict_used_overtime: dict,
                              weekday: dict, LIMIT: float = 36.0) -> dict:
    """只计算标准人员的 36 小时截断
    按 weekday 分组（节假日/休息日/工作日），每组内按小时数降序排列
    节假日优先 fill：节假日总 ≥ 36 时按降序逐个扣直到 36
    否则节假日全 pay，剩余预算给休息日，最后给工作日
    """
    for id in valid_ids:
        if not dict_used_overtime[id]:
            continue

        day_dict = dict_used_overtime[id]

        # 单次遍历分组：按 weekday 分成 3 个子字典
        holiday_items, weekday_items, workday_items = {}, {}, {}
        for d, v in day_dict.items():
            wd = weekday.get(d)
            if wd == 3:
                holiday_items[d] = v[0]
            elif wd == 2:
                weekday_items[d] = v[0]
            elif wd == 1.5:
                workday_items[d] = v[0]

        dictfee_holiday = dict(
            sorted(holiday_items.items(), key=lambda item: item[1], reverse=True)
        )
        dictfee_weekday = dict(
            sorted(weekday_items.items(), key=lambda item: item[1], reverse=True)
        )
        dictfee_workday = dict(
            sorted(workday_items.items(), key=lambda item: item[1], reverse=True)
        )

        cash_dict = {}
        rest_dict = {}
        remainer = LIMIT

        def assign_all(category_dict):
            """整组全 pay，剩余预算相应减少"""
            nonlocal remainer, cash_dict
            for date_key, hours in category_dict.items():
                cash_dict[date_key] = hours
            remainer = round(remainer - floor_half(sum(category_dict.values())), 2)

        def fill_until_zero(category_dict):
            """按降序逐个扣，直到扣满 36（remainer ≤ 0）或该组扣完"""
            nonlocal remainer, cash_dict, rest_dict
            for date_key, hours in category_dict.items():
                if remainer <= 0:
                    rest_dict[date_key] = hours
                    continue
                if hours <= remainer:
                    cash_dict[date_key] = hours
                    remainer = round(remainer - hours, 2)
                else:
                    cash_dict[date_key] = remainer
                    rest_dict[date_key] = round(hours - remainer, 2)
                    remainer = 0
            return remainer == 0

        if sum(dictfee_holiday.values()) >= remainer:
            fill_until_zero(dictfee_holiday)
        else:
            assign_all(dictfee_holiday)
            if remainer > 0:
                if sum(dictfee_weekday.values()) >= remainer:
                    fill_until_zero(dictfee_weekday)
                else:
                    assign_all(dictfee_weekday)
                    if remainer > 0:
                        fill_until_zero(dictfee_workday)

        for d in dict_used_overtime[id]:
            hours, night_type = (
                dict_used_overtime[id][d][0],
                dict_used_overtime[id][d][1],
            )
            pay = cash_dict.get(d, 0)
            comp_off = rest_dict.get(d, 0)
            if d not in cash_dict and d not in rest_dict:
                pay = 0
                comp_off = hours
            dict_used_overtime[id][d] = [hours, night_type, pay, comp_off]

    return dict_used_overtime


def calculate_night_truncation(night_ids: list, dict_used_overtime_night: dict,
                               weekday: dict,
                               dict_overtime_night_original: dict = None) -> dict:
    """夜班加班截断计算
    按系数从大到小遍历：
    Phase 1 (加班费): sum < pay_limit → x=t, y=0
    Phase 2 (截断): 达到 pay_limit 时, x=剩余, y=t-x
    Phase 3 (串休): sum_comp < comp_limit → x=0, y=t
    Phase 4 (串休截断): 达到 comp_limit 时, y=剩余, 超出部分 x=0,y=0

    Args:
        night_ids: 夜班人员工号列表
        dict_used_overtime_night: 夜班加班数据字典
        weekday: 工作日类型字典
        dict_overtime_night_original: 调整前的夜班加班数据（用于计算total_hours）
    Returns:
        处理后的夜班加班数据字典
    """
    for emp_id in night_ids:
        entries = dict_used_overtime_night[emp_id]
        if not entries:
            continue

        # 优先使用调整前的总加班（不包含adjustments的影响）
        if dict_overtime_night_original and emp_id in dict_overtime_night_original:
            total_hours = sum(
                day + night
                for _, (day, night), _, _ in dict_overtime_night_original[emp_id]
            )
        else:
            total_hours = sum(x[1] for x in entries)
        work_hours = sum([1 for x in weekday if weekday[x] == 1.5]) * 8

        if total_hours > work_hours:
            overtime = total_hours - work_hours
            pay_limit = min(overtime, 36)
            comp_limit = max(overtime - 36, 0)
        else:
            pay_limit = 0
            comp_limit = 0

        sum_pay = 0.0
        sum_comp = 0.0

        for i, entry in enumerate(entries):
            date, t, coe, night_type = entry
            x, y = 0.0, 0.0
            remaining = t

            if sum_pay < pay_limit and remaining > 0:
                pay_amount = min(remaining, pay_limit - sum_pay)
                x = pay_amount
                sum_pay += pay_amount
                remaining -= pay_amount

            if comp_limit > 0 and sum_comp < comp_limit and remaining > 0:
                comp_amount = min(remaining, comp_limit - sum_comp)
                y = comp_amount
                sum_comp += comp_amount
                remaining -= comp_amount

            entries[i] = [date, t, coe, night_type, x, y]

        dict_used_overtime_night[emp_id] = entries
        logging.info(f"夜班人员 {emp_id}: 加班费 {sum_pay:.2f}h, 串休 {sum_comp:.2f}h")
    return dict_used_overtime_night


def calculate_main(ids: list, night_ids: list = None) -> None:
    """主计算函数
    Args:
        ids: 标准人员工号列表
        night_ids: 夜班人员工号列表
    """
    if night_ids is None:
        night_ids = []

    global dict_id_to_name
    dict_id_to_name = {}

    logging.info("开始读取进出场记录...")
    wb = load_workbook_safe("进出场记录.xlsx")
    dict_in_out = {}
    ws = wb["Sheet1"]
    for row in ws.rows:
        if row[17].value == "进厂":
            emp_id_value = row[11].value
            emp_name_value = row[1].value
            if emp_id_value is not None and emp_name_value is not None:
                dict_id_to_name[emp_id_value] = emp_name_value
            dict_in_out.setdefault(emp_id_value, []).append((row[0].value, "in"))
        elif row[17].value == "出厂":
            emp_id_value = row[11].value
            emp_name_value = row[1].value
            if emp_id_value is not None and emp_name_value is not None:
                dict_id_to_name[emp_id_value] = emp_name_value
            dict_in_out.setdefault(emp_id_value, []).append((row[0].value, "out"))

    logging.info("开始转换数据格式...")
    for k in dict_in_out:
        temp = {}
        for v in dict_in_out[k]:
            d, _, t = str(v[0]).partition(" ")
            temp.setdefault(format_date(d), []).append((t, v[1]))
        dict_in_out[k] = {
            d: dedup_consecutive(sorted(temp[d], key=lambda x: x[0]))
            for d in sorted(temp)
        }
    dict_in_out = {k: dict_in_out[k] for k in sorted(dict_in_out)}
    logging.info("开始计算加班时长...")
    dict_overtime = {}
    dict_overtime_night = {}
    for id in ids:
        if id not in dict_in_out:
            logging.warning(f"工号 {id} 不存在于进出场记录中，已跳过")
            continue
        dict_overtime[id] = {}
        dict_overtime_night[id] = []
        for d in dict_in_out[id]:
            if id in night_ids:
                dict_overtime_night[id].append(
                    calculate_dict_overtime_night(dict_in_out, id, d)
                )
            else:
                dict_overtime[id][d] = calculate_dict_overtime(dict_in_out, id, d)
    
    dict_used_overtime = deepcopy(dict_overtime)
    for id in night_ids:
        dict_used_overtime.pop(id, None)

    # 保存夜班原始加班数据（用于统计表"时长"列和明细表E-H列）
    dict_overtime_night_original = deepcopy(dict_overtime_night)

    logging.info("开始应用调整值...")
    wb2 = load_workbook_safe("计算结果.xlsx")
    ws2 = wb2["统计表"]
    rows = list(ws2.rows)
    total_chuanxiu = apply_row10_adjustments(
        rows,
        dict_used_overtime,
        dict_overtime,
        dict_overtime_night,
        night_ids,
    )
    logging.info(f"调整值应用完成，共处理 {len(total_chuanxiu)} 名员工")

    logging.info("开始处理夜班人员数据...")
    dict_used_overtime_night = {}
    for id in night_ids:
        dict_used_overtime_night[id] = []
        for data in dict_overtime_night[id]:
            result1 = [x[0] if isinstance(x, tuple) else x for x in data]
            result2 = [x[1] if isinstance(x, tuple) else x for x in data]
            dict_used_overtime_night[id].append(result1)
            dict_used_overtime_night[id].append(result2)
        dict_used_overtime_night[id] = [
            x for x in dict_used_overtime_night[id] if x[1] != 0
        ]
        dict_used_overtime_night[id].sort(key=lambda x: x[2], reverse=True)

    logging.info("开始计算36小时截断...")
    valid_ids = [id for id in ids if id not in night_ids]
    dict_used_overtime = calculate_36h_truncation(
        valid_ids, dict_used_overtime, weekday
    )
    dict_used_overtime_night = calculate_night_truncation(
        night_ids, dict_used_overtime_night, weekday, dict_overtime_night_original
    )

    logging.info("开始写入Excel结果...")
    write_dict_overtime_to_excel(
        valid_ids,
        night_ids,
        dict_in_out,
        dict_overtime,
        dict_overtime_night,
        dict_overtime_night_original,
        dict_used_overtime,
        dict_used_overtime_night,
        total_chuanxiu
    )
    logging.info("计算完成，结果已写入 计算结果.xlsx")


# if __name__ == "__main__":
#     logging.info("===== 加班计算程序启动 =====")
#     calculate_main(["Q4642"], night_ids=["Q4642"])
