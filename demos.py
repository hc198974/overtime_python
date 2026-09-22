# -*- coding: utf-8 -*-
import calendar
import datetime
import json
import config
import tkinter
import tkinter.simpledialog
import requests
from lxml import etree
import win32com.client
import time
import functools
import os
import logging
from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def run_time(fn):  # 用于测试方法运行时间的装饰器
    @functools.wraps(fn)
    def wrapper(*args, **kw):
        start = time.perf_counter()
        res = fn(*args, **kw)
        logging.info('%s 运行了 %f 秒', fn.__name__, time.perf_counter() - start)
        return res
    return wrapper


class Crili(object):
    """
    万年日历接口数据抓取
    Params:year 四位数年份字符串
    """

    CACHE_TTL_SECONDS = 24 * 3600

    def __init__(self, year, month):
        self.year = year
        self.month = month

    def _get_cache_path(self):
        """返回本地缓存文件路径。"""
        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
        os.makedirs(cache_dir, exist_ok=True)
        return os.path.join(cache_dir, f"holiday_{self.year}_{self.month}.json")

    def _load_cache(self, allow_stale=False):
        """加载本地缓存：未过期直接返回，过期时仅 allow_stale=True 才返回旧值。"""
        cache_path = self._get_cache_path()
        if not os.path.exists(cache_path):
            return None

        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            if not isinstance(payload, dict):
                return None
            fetched_at = payload.get("fetched_at")
            result = payload.get("result")
            if not fetched_at or not isinstance(result, dict):
                return None

            dt = datetime.datetime.fromisoformat(fetched_at)
            age_seconds = (datetime.datetime.now() - dt).total_seconds()
            if age_seconds <= self.CACHE_TTL_SECONDS:
                return result
            if allow_stale:
                return result
            return None
        except (ValueError, TypeError, OSError, json.JSONDecodeError):
            logging.warning("读取节假日缓存失败，忽略旧缓存：%s", cache_path)
            return None

    def _save_cache(self, result):
        """将当前 result 持久化到本地缓存。"""
        cache_path = self._get_cache_path()
        payload = {
            "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "result": result,
        }
        try:
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            logging.info("节假日结果已缓存到本地：%s", cache_path)
        except OSError as exc:
            logging.warning("写入节假日缓存失败：%s", exc)

    def _fetch_holiday_dates_from_api(self):
        """备用接口：从 timor.tech 获取法定节假日日期，返回 set("YYYY-MM-DD")。"""
        url = f"https://timor.tech/api/holiday/year/{self.year}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
        }
        try:
            response = requests.get(url, headers=headers, timeout=20)
            response.raise_for_status()
            data = response.json()
            holiday_set = set()
            for date_str, info in data.get("holiday", {}).items():
                if isinstance(info, dict) and info.get("wage") == 3:
                    holiday_set.add(date_str)
            return holiday_set
        except Exception as exc:
            logging.warning("备用节假日接口请求失败：%s", exc)
            return set()

    def parseHTML(self):
        """页面解析：优先使用本地缓存，未过期时不联网；过期后才联网刷新并写回缓存。"""
        global weekday

        cached_result = self._load_cache(allow_stale=False)
        if cached_result is not None:
            weekday = cached_result
            logging.info("使用本地节假日缓存，跳过联网请求：%s", self._get_cache_path())
            return cached_result

        result = {}
        # 2026年法定节假日列表，改年份同样得改
        holiday_3x = [
            "20260101",
            "20260215",
            "20260216",
            "20260217",
            "20260218",
            "20260405",
            "20260501",
            "20260502",
            "20260619",
            "20260925",
            "20261001",
            "20261002",
            "20261003"
        ]

        c = calendar.monthrange(self.year, self.month)[1]
        for day in range(1, c + 1):
            temp = datetime.datetime(self.year, self.month, day)
            result[temp.strftime("%Y%m%d")] = 2 if temp.weekday() > 4 else 1.5

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

        try:
            payload = {"q": f"{self.year}-{self.month}"}
            response = requests.get(url, headers=headers, params=payload, timeout=20)
            response.raise_for_status()
            element = etree.HTML(response.text)
            html = element.xpath('//div[@class="wnrl_riqi"]')

            if html and len(html) >= c:
                # 获取节点属性，优先采用万年历原始判断
                for i in range(c):
                    item = html[i].xpath("./a")[0].attrib
                    if item.get("id") == "wnrl_riqi_id_" + str(i):
                        temp = datetime.datetime(self.year, self.month, i + 1)
                        day_key = temp.strftime("%Y%m%d")
                        if "class" in item:
                            cls = item["class"]
                            if cls in ("wnrl_riqi_xiu", "wnrl_riqi_mo"):
                                result[day_key] = 2
                            elif cls == "wnrl_riqi_ban":
                                result[day_key] = 1.5
                        else:
                            result[day_key] = 2 if temp.weekday() > 4 else 1.5
            else:
                raise ValueError("万年历返回数据为空或不完整")
        except Exception as exc:
            logging.warning("万年历接口不可用，自动切换到备用接口：%s", exc)

        holiday_dates = self._fetch_holiday_dates_from_api()
        if holiday_dates:
            for day_key in list(result):
                d = datetime.datetime.strptime(day_key, "%Y%m%d").strftime("%Y-%m-%d")
                if d in holiday_dates:
                    result[day_key] = 3
        else:
            # 保持原有的硬编码兜底，确保 2026 年节假日也能正确识别
            result.update({k: 3 for k in holiday_3x if k in result})

        if result:
            self._save_cache(result)

        weekday = result
        return result


class Cwindow(object):
    def __init__(self):
        self.year = config.YEAR
        self.month = config.MONTH
        self.start_calculation = False

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
            initialvalue=config.MONTH,
        )
        self.month = month

    def dealSheet(self):
        # 重置并生成统计表：A=姓名 B=职号 C=日报日期（从月初到最后一天）
        from openpyxl import load_workbook

        path = "计算结果.xlsx"
        wb = load_workbook(path)

        # 1. 从记录表提取 姓名 -> 职号（数据从第 3 行起，第1列姓名、第2列职号）
        ws_jilu = wb["记录表"]
        emp_list = []
        for r in range(3, ws_jilu.max_row + 1):
            name = ws_jilu.cell(row=r, column=1).value
            emp_id = ws_jilu.cell(row=r, column=2).value
            if name and emp_id:
                emp_list.append((str(name).strip(), str(emp_id).strip()))

        # 2. 清空统计表中除题头（第 1 行）以外的数据
        ws_tongji = wb["统计表"]
        if ws_tongji.max_row >= 2:
            ws_tongji.delete_rows(2, ws_tongji.max_row - 1)

        # 3. 按全局年月，为每个人生成从月初到最后一天的逐日记录（C 列日期格式 YYYY-MM-DD）
        year, month = config.YEAR, config.MONTH
        _, last_day = calendar.monthrange(year, month)
        row = 2
        for name, emp_id in emp_list:
            for day in range(1, last_day + 1):
                ws_tongji.cell(row=row, column=1, value=name)
                ws_tongji.cell(row=row, column=2, value=emp_id)
                ws_tongji.cell(row=row, column=3, value=f"{year}-{month:02d}-{day:02d}")
                row += 1

        wb.save(path)
        logging.info("统计表已重置：%d 人 × %d 天（%d-%02d）共写入 %d 行。",
                     len(emp_list), last_day, year, month, row - 2)

    def shutDown(self):
        self.start_calculation = True
        root.destroy()

    def on_close(self):
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
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        # 加入消息循环
        root.mainloop()
        return self.start_calculation


# 这是用来获得当月节假日（wage=3）的类，接口来自 timor.tech
class Ccal:
    def __init__(self, year, month):
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
                if int(date.split("-")[0]) == self.month and info["wage"] == 3:
                    holidays.append(date)
            return holidays
        except requests.RequestException as e:
            logging.error("请求出错: %s", e)
        except (KeyError, ValueError) as e:
            logging.error("解析数据出错: %s", e)
        return []
