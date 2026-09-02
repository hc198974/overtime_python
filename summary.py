"""
LangChain 1.x 推荐写法（自包含版）：从 `计算结果.xlsx` 中提取指定人员的加班数据并对其进行总结。

特性：
- 使用 `langchain-openai` 的 `ChatOpenAI` 作为推理后端（OpenAI 兼容接口）；
- 支持两种后端，通过 --backend / 环境变量 LLM_BACKEND 切换，默认使用本地 llama.cpp：
    * llama    ：本地 llama.cpp server（OpenAI 兼容，默认 http://localhost:8080/v1）
    * deepseek ：远程 DeepSeek 服务（默认 https://api.deepseek.com/v1）
- 使用 `langchain.agents.create_agent` 构建标准 ReAct Agent，工具调用循环由框架托管，
  不再手写多轮 ToolMessage 往返，也不再需要自定义 BaseChatModel 子类。
- 所有 Excel 读取逻辑与工具均内置，不依赖其它文件。

数据约定：
- 明细表列：总加班(5) / 工作日(6) / 休息日(7) / 节假日(8) / I列(9)=串休小时数（本月已扣除串休数）
  / 转加班费小时(10) / 工作日(11) / 休息日(12) / 节假日(13) / 加班费(15) / 工作日(16) / 休息日(17) / 节假日(18)
- 可串休时长来自【记录表最后一列】（“可串休时间”）；
- 其余汇总来自明细表。

配置统一放在 .env（均可选）：

  # 后端选择：llama（默认）或 deepseek
  LLM_BACKEND=llama

  # 本地 llama.cpp 配置
  LLAMACPP_BASE_URL=http://localhost:8080/v1
  LLAMACPP_MODEL=local-model
  LLAMACPP_API_KEY=not-needed
  LLAMACPP_TIMEOUT=600

  # 远程 DeepSeek 配置
  DEEPSEEK_BASE_URL=https://api.deepseek.com
  DEEPSEEK_MODEL=deepseek-v4-flash
  DEEPSEEK_API_KEY=your-key
  DEEPSEEK_TIMEOUT=120
"""

import os
import datetime
import argparse
from typing import Dict, Any, Optional, List

from openpyxl import load_workbook
from langchain_core.tools import tool
from langchain_core.messages import HumanMessage


# ---------------------------------------------------------------------------
# 加载 .env
# ---------------------------------------------------------------------------
def _load_dotenv(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key, val = key.strip(), val.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = val


_load_dotenv()

# --- 本地 llama.cpp 配置（默认后端） ---
LLAMACPP_BASE_URL = os.environ.get("LLAMACPP_BASE_URL", "http://localhost:8080/v1")
LLAMACPP_MODEL = os.environ.get("LLAMACPP_MODEL", "local-model")
LLAMACPP_API_KEY = os.environ.get("LLAMACPP_API_KEY", "not-needed")
LLAMACPP_TIMEOUT = int(os.environ.get("LLAMACPP_TIMEOUT", "600"))

# --- 远程 DeepSeek 配置 ---
DEEPSEEK_BASE_URL = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
DEEPSEEK_MODEL = os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash")
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
DEEPSEEK_TIMEOUT = int(os.environ.get("DEEPSEEK_TIMEOUT", "120"))

# 默认后端：llama（本地）或 deepseek（远程）
DEFAULT_BACKEND = os.environ.get("LLM_BACKEND", "llama").lower()
if DEFAULT_BACKEND not in ("llama", "deepseek"):
    DEFAULT_BACKEND = "llama"

EXCEL_PATH = "计算结果.xlsx"


def _backend_config(backend: str):
    """返回 (base_url, model, api_key, timeout) 元组。"""
    if backend == "deepseek":
        return DEEPSEEK_BASE_URL, DEEPSEEK_MODEL, DEEPSEEK_API_KEY, DEEPSEEK_TIMEOUT
    return LLAMACPP_BASE_URL, LLAMACPP_MODEL, LLAMACPP_API_KEY, LLAMACPP_TIMEOUT


# ---------------------------------------------------------------------------
# Excel 读取
# ---------------------------------------------------------------------------
def _load_excel(path: str = EXCEL_PATH):
    """加载计算结果 Excel，返回 (wb, ws_mingxi, ws_jilu, date_col)。"""
    if not os.path.exists(path):
        raise FileNotFoundError(f"文件不存在：{path}")
    wb = load_workbook(path, data_only=True)
    try:
        ws_mingxi = wb["明细表"]
        ws_jilu = wb["记录表"]
    except KeyError as e:
        wb.close()
        raise KeyError(f"缺少工作表：{e}")

    date_col: Dict[str, int] = {}
    for cell in ws_jilu[2]:
        if isinstance(cell.value, datetime.datetime):
            date_col[cell.value.strftime("%Y%m%d")] = cell.column
    return wb, ws_mingxi, ws_jilu, date_col


def _resolve_emp_id(ws_mingxi, key: str) -> Optional[str]:
    """按工号精确匹配；否则按姓名子串匹配，返回工号。"""
    target = str(key).strip()
    for row in ws_mingxi.iter_rows(min_row=4):
        if row[2].value is not None and str(row[2].value).strip() == target:
            return target
    for row in ws_mingxi.iter_rows(min_row=4):
        name = row[1].value
        if name is not None and (target in str(name) or str(name) in target):
            return str(row[2].value).strip()
    return None


def _extract_summary(ws_mingxi, ws_jilu, emp_id: str) -> Optional[Dict[str, Any]]:
    """提取某员工的总览数据（来自明细表；可串休来自记录表末列）。"""
    target = str(emp_id).strip()
    resolved = _resolve_emp_id(ws_mingxi, target)
    if resolved is None:
        return None
    for row in ws_mingxi.iter_rows(min_row=4):
        if row[2].value is None:
            continue
        if str(row[2].value).strip() == resolved:
            total = row[4].value or 0
            work = row[5].value or 0
            rest = row[6].value or 0
            holiday = row[7].value or 0
            deducted_comp_off = row[8].value or 0          # I 列（第9列）：串休小时数 = 本月已扣除串休数
            pay_total_hours = row[9].value or 0
            pay_work_hours = row[10].value or 0
            pay_rest_hours = row[11].value or 0
            pay_holiday_hours = row[12].value or 0
            money_total = row[14].value or 0
            money_p = row[15].value or 0
            money_q = row[16].value or 0
            money_r = row[17].value or 0
            return {
                "emp_id": resolved,
                "name": str(row[1].value) if row[1].value is not None else "",
                "total_hours": float(total),
                "work_hours": float(work),
                "rest_hours": float(rest),
                "holiday_hours": float(holiday),
                "deducted_comp_off": float(deducted_comp_off),
                "pay_total_hours": float(pay_total_hours),
                # 可串休时长来自记录表最后一列（“可串休时间”），而非明细表推算
                "comp_off_hours": _extract_comp_off(ws_jilu, resolved),
                "pay_work_hours": float(pay_work_hours),
                "pay_rest_hours": float(pay_rest_hours),
                "pay_holiday_hours": float(pay_holiday_hours),
                "money_total": float(money_total),
                "money_p": float(money_p),
                "money_q": float(money_q),
                "money_r": float(money_r),
            }
    return None


def _extract_comp_off(ws_jilu, emp_id: str) -> Optional[float]:
    """从记录表【最后一列】提取指定工号的“可串休时间”。
    记录表 emp_id 位于第 2 列；最后一列表头为“可串休时间”。若单元格为空则返回 None。
    """
    target = str(emp_id).strip()
    last_col = ws_jilu.max_column
    for row in ws_jilu.iter_rows(min_row=3, max_col=last_col, values_only=True):
        if row[1] is None:
            continue
        if str(row[1]).strip() == target:
            val = row[last_col - 1]
            if val is None or val == "":
                return None
            try:
                return float(val)
            except (ValueError, TypeError):
                return None
    return None


def _extract_daily(ws_jilu, date_col: Dict[str, int], emp_id: str) -> Dict[str, float]:
    """从记录表提取指定工号的按日期加班值，返回 {yyyymmdd: value}。

    记录表 emp_id 位于第 2 列；第 2 行为日期头（datetime），date_col 提供 日期->列号 映射。
    对匹配到的员工行，逐列读取各日期对应的加班时长。
    """
    target = str(emp_id).strip()
    for row in ws_jilu.iter_rows(min_row=3):
        if row[1].value is None:
            continue
        if str(row[1].value).strip() == target:
            res: Dict[str, float] = {}
            for d, col in date_col.items():
                v = ws_jilu.cell(row=row[0].row, column=col).value
                if v is None or v == "":
                    continue
                try:
                    res[d] = float(v)
                except (ValueError, TypeError):
                    pass
            return res
    return {}


# ---------------------------------------------------------------------------
# 工具（@tool）
# ---------------------------------------------------------------------------
@tool
def list_employee_ids() -> str:
    """列出明细表中所有员工的工号。"""
    wb, ws_mingxi, _, _ = _load_excel()
    try:
        ids = []
        for row in ws_mingxi.iter_rows(min_row=4):
            v = row[2].value
            if v is not None and str(v).strip():
                ids.append(str(v).strip())
    finally:
        wb.close()
    if not ids:
        return "明细表中未找到任何工号。"
    return "可用工号：" + "、".join(ids)


@tool
def find_employee_id(name: str) -> str:
    """根据姓名（模糊匹配）在明细表中查找对应的工号。参数 name：员工姓名片段，如 '曲'。"""
    wb, ws_mingxi, _, _ = _load_excel()
    try:
        matches = []
        for row in ws_mingxi.iter_rows(min_row=4):
            v = row[1].value
            if v is not None and name in str(v):
                matches.append(f"{str(row[2].value).strip()}({str(v).strip()})")
    finally:
        wb.close()
    if not matches:
        return f"未找到包含 '{name}' 的员工。"
    return "匹配到的员工：" + "、".join(matches)


@tool
def get_employee_summary(emp_id: str) -> str:
    """根据工号或姓名从明细表中提取该人员的总体加班汇总数据（总时长、工作日/休息日/节假日小时、
    转为加班费小时、可串休小时、本月已扣除串休数、加班费金额等），返回结构化的中文文本描述。
    参数 emp_id：员工的工号或姓名，例如 'Q6007' 或 '曲书成'。"""
    wb, ws_mingxi, ws_jilu, _ = _load_excel()
    try:
        s = _extract_summary(ws_mingxi, ws_jilu, emp_id)
    finally:
        wb.close()
    if not s:
        return f"在明细表中未找到与 '{emp_id}' 匹配的员工（工号或姓名）。"
    name_part = f"（{s['name']}）" if s.get("name") else ""
    lines = [
        f"工号 {s['emp_id']}{name_part} 的总体加班汇总：",
        f"- 总加班时长：{s['total_hours']:.2f} 小时",
        f"- 工作日加班：{s['work_hours']:.2f} 小时",
        f"- 休息日加班：{s['rest_hours']:.2f} 小时",
        f"- 节假日加班：{s['holiday_hours']:.2f} 小时",
        f"- 转为加班费小时（总）：{s['pay_total_hours']:.2f} 小时",
        (f"- 可串休小时：{s['comp_off_hours']:.2f} 小时"
         if s["comp_off_hours"] is not None
         else "- 可串休小时：未记录（记录表“可串休时间”列为空）"),
        f"- 本月已扣除串休数（明细表 I 列“串休小时数”）：{s['deducted_comp_off']:.2f} 小时",
        f"- 转为加班费（金额）：{s['money_total']:.2f} 元",
        f"- 其中：工作日加班费 {s['money_p']:.2f} 元 / 休息日 {s['money_q']:.2f} 元 / 节假日 {s['money_r']:.2f} 元",
    ]
    return "\n".join(lines)


@tool
def get_employee_daily(emp_id: str) -> str:
    """根据工号或姓名从记录表中提取该人员按日期的加班明细，返回中文文本（含表格）。
    参数 emp_id：员工的工号或姓名，例如 'Q6007' 或 '曲书成'。"""
    wb, ws_mingxi, ws_jilu, date_col = _load_excel()
    try:
        recs = _extract_daily(ws_jilu, date_col, emp_id)
    finally:
        wb.close()
    if not recs:
        return f"未找到 {emp_id} 的每日加班记录。"
    total = sum(recs.values())
    lines = [
        f"工号 {emp_id} 的每日加班明细（共 {len(recs)} 天，合计 {total:.2f} 小时）：",
        "| 日期 | 加班小时 |",
        "|------|--------|",
    ]
    for d in sorted(recs):
        lines.append(f"| {d[:4]}-{d[4:6]}-{d[6:8]} | {recs[d]:.2f} |")
    return "\n".join(lines)


TOOLS = [list_employee_ids, find_employee_id, get_employee_summary, get_employee_daily]


SYSTEM_PROMPT = """你是一个人力资源加班数据分析助手。你可以调用工具从《计算结果.xlsx》中查询某位员工的加班数据。

工作流程：
1. 用户可能给出工号（如 Q6007）或姓名（如 曲书成）。get_employee_summary 与 get_employee_daily 需要工号或姓名；
   若只给姓名，先用 find_employee_id 模糊查找工号（结果唯一时取之），再继续；若直接给了工号可跳过此步。
2. 调用 get_employee_summary 获取总体汇总（总加班、工作日/休息日/节假日构成、转加班费、可串休、本月已扣除串休数、加班费金额等）。
3. 必要时调用 get_employee_daily 获取每日明细。
4. 综合上述数据，用简洁、正式的中文对该员工的加班情况做出总结，至少包含：总加班量、工作日/休息日/节假日构成、加班费与串休情况，不用给出建议。
5. 输出使用 Markdown，标题层级清晰；若数据缺失（如可串休小时未记录）请如实说明。"""


# ---------------------------------------------------------------------------
# 1.x 推荐：ChatOpenAI + langchain.agents.create_agent（可切换后端）
# ---------------------------------------------------------------------------
def _to_openai_base_url(url: str) -> str:
    """将用户配置地址规整为 ChatOpenAI 所需的 base_url（需指向 .../v1）。

    ChatOpenAI 会在 base_url 后追加 /chat/completions，因此这里保留到 /v1：
      - http://localhost:8080          -> http://localhost:8080/v1
      - https://api.deepseek.com        -> https://api.deepseek.com/v1
      - .../v1/chat/completions          -> 去掉末尾 /chat/completions
    """
    u = (url or "").strip().rstrip("/")
    if not u:
        raise ValueError("base_url 不能为空，请在 .env 中配置对应后端的地址。")
    if u.endswith("/chat/completions"):
        return u[: -len("/chat/completions")]
    if u.endswith("/v1"):
        return u
    return u + "/v1"


def build_llm(
    backend: str = DEFAULT_BACKEND,
    model: Optional[str] = None,
    base_url: Optional[str] = None,
    api_key: Optional[str] = None,
    timeout: Optional[int] = None,
    temperature: float = 0.0,
    max_tokens: int = 2048,
):
    """构建一个 1.x 标准的 ChatOpenAI 实例，指向指定后端的 OpenAI 兼容端点。

    backend 可为 "llama"（默认，本地）或 "deepseek"（远程）；
    任一连接参数若未显式传入，则回退到对应后端的默认配置。
    """
    if backend not in ("llama", "deepseek"):
        raise ValueError(f"未知后端：{backend!r}，仅支持 'llama' 或 'deepseek'。")
    def_bu, def_model, def_key, def_to = _backend_config(backend)

    base_url = base_url or def_bu
    model = model or def_model
    api_key = api_key or def_key
    timeout = timeout or def_to

    from langchain_openai import ChatOpenAI

    return ChatOpenAI(
        model=model,
        base_url=_to_openai_base_url(base_url),
        api_key=api_key,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=timeout,
        max_retries=2,
    )


def run_agent(
    question: str,
    backend: str = DEFAULT_BACKEND,
    model: Optional[str] = None,
    base_url: Optional[str] = None,
    api_key: Optional[str] = None,
    timeout: Optional[int] = None,
    temperature: float = 0.0,
    max_tokens: int = 2048,
) -> str:
    """运行 ReAct Agent 回答提问，返回最终中文总结文本。

    关键点（1.x 推荐）：
      1. bind 工具交给框架：create_agent(model, tools, system_prompt=...) 内部已调用 model.bind_tools(tools)；
      2. 工具调用循环由 langchain.agents 托管，无需手写 _generate + ToolMessage 多轮；
      3. 输入为 {"messages": [...]} 标准消息列表，输出也是消息列表。
    """
    from langchain.agents import create_agent

    llm = build_llm(backend, model, base_url, api_key, timeout, temperature, max_tokens)
    agent = create_agent(llm, TOOLS, system_prompt=SYSTEM_PROMPT)

    result = agent.invoke({"messages": [HumanMessage(content=question)]})

    messages = result.get("messages", [])
    # 取最后一条带正文的消息作为最终答案（兜底处理）
    answer = ""
    for m in reversed(messages):
        if getattr(m, "content", None):
            answer = m.content
            break
    return answer


def summarize_employee(
    emp_id: str,
    backend: str = DEFAULT_BACKEND,
    model: Optional[str] = None,
    base_url: Optional[str] = None,
    api_key: Optional[str] = None,
    timeout: Optional[int] = None,
) -> str:
    """便捷入口：直接查询并总结某个工号/姓名的加班数据。"""
    return run_agent(
        f"请提取工号 {emp_id} 的加班数据并做总结。",
        backend=backend,
        model=model,
        base_url=base_url,
        api_key=api_key,
        timeout=timeout,
    )


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="LangChain 1.x 推荐写法：从 计算结果.xlsx 提取某人加班数据并总结（支持 llama.cpp / deepseek）"
    )
    parser.add_argument("--id", help="员工工号或姓名，例如 Q6007 / 曲书成")
    parser.add_argument("--question", help="自由提问（与 --id 二选一）")
    parser.add_argument("--backend", default=DEFAULT_BACKEND,
                        choices=["llama", "deepseek"],
                        help="推理后端：llama（本地，默认）或 deepseek（远程）")
    parser.add_argument("--model", default=None, help="模型名（覆盖后端默认值）")
    parser.add_argument("--base-url", default=None, help="OpenAI 兼容接口地址（覆盖后端默认值）")
    parser.add_argument("--api-key", default=None, help="API Key（覆盖后端默认值）")
    parser.add_argument("--timeout", type=int, default=None, help="请求超时秒数（覆盖后端默认值）")
    args = parser.parse_args()

    if args.id and not args.question:
        args.question = f"请提取工号 {args.id} 的加班数据并做总结。"
    if not args.question:
        parser.error("请提供 --id 或 --question")

    try:
        answer = run_agent(
            args.question,
            backend=args.backend,
            model=args.model,
            base_url=args.base_url,
            api_key=args.api_key,
            timeout=args.timeout,
        )
    except Exception as e:  # noqa: BLE001
        hint = (
            "请确认本地 llama.cpp server 已启动（默认 localhost:8080）且模型支持 function calling；"
            if args.backend == "llama"
            else "请确认 .env 中已配置 DEEPSEEK_API_KEY 且网络可访问 DeepSeek 接口，且模型支持 function calling。"
        )
        print(f"[错误] Agent 运行失败：{e}\n{hint}")
        return

    print(answer)

    # 写入报告（与历史行为一致）
    out_dir = "reporter"
    os.makedirs(out_dir, exist_ok=True)
    name = args.id or "query"
    out_path = os.path.join(out_dir, f"report_{name}.md")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(answer)
    print(f"\n报告已写入 {out_path}")


if __name__ == "__main__":
    main()
