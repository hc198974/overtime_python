# Overtime 加班 / 串休管理工具

基于 Python 的加班统计与串休管理工具。结合 Excel 进出场打卡记录、万年历节假日数据，自动计算每位员工的加班时长与加班费，并提供串休抵扣操作与 AI 数据总结。

## 功能概览

- **加班计算**：读取《进出场记录.xlsx》的进出厂打卡记录，区分「标准人员」与「夜班人员」，按工作日 / 公休日 / 节假日倍率计算每日加班时长，写回《计算结果.xlsx》。
- **节假日倍率**：通过 `Crili` 从万年历网站抓取指定年月的日期类型（工作日 1.5 / 公休日 2 / 节假日 3），并叠加 2026 年法定节假日。
- **统计表骨架生成**：`dealSheet` 一键清空统计表并按「全局年月」为每位员工生成「姓名 / 职号 / 逐日日期」骨架，供后续计算填充。
- **串休抵扣**：`huanxiu.py` 图形界面，输入职号（自动带出姓名）、年月、串休时长，按「工作日 → 公休日 → 节假日」顺序抵扣统计表 K 列；支持清空指定职号+年月的 K 列。
- **AI 总结**：`summary.py` 使用 LangChain ReAct Agent + LLM（本地 llama.cpp 或远程 DeepSeek），从《计算结果.xlsx》提取指定员工加班数据并生成中文 Markdown 报告。

## 目录结构

```
overtime2/
├── config.py              # 全局年月配置（手动修改的唯一入口）
├── main.py                # 主入口：弹窗 → 读取记录表职号 → 调用 calculate_main 计算
├── process_attendance.py  # 核心计算模块：进出场记录 → 加班时长 → 写回计算结果.xlsx
├── demos.py               # 工具类：Crili(万年历) / Cwindow(月份窗口+dealSheet) / Ccal(节假日接口)
├── huanxiu.py             # 串休扣减图形界面（tkinter）
├── summary.py             # AI 加班总结（LangChain + llama.cpp / DeepSeek）
├── requirements.txt       # 依赖清单
├── 进出场记录.xlsx         # 源数据：员工进出厂打卡记录（需手动提供/导出）
├── 计算结果.xlsx           # 计算结果：记录表 / 统计表 / 明细表 / 汇总表
└── reporter/              # summary.py 生成的 Markdown 报告输出目录
```

> 说明：旧版 README 提到的 `overtime.py`、`原始数据.xlsm`、`Cmacro` 等已不在当前代码中，请以本目录实际文件为准。

## 环境依赖与安装

建议使用 **Python 3.9+**（LangChain 1.x 系列的最低要求）。在 Windows 下运行（GUI 与 `win32com` 依赖 Windows）。

1. 安装依赖：

   ```bash
   pip install -r requirements.txt
   ```

2. 运行 `summary.py` 额外需要 `langchain-openai`（当前 `requirements.txt` 未列出），请补装：

   ```bash
   pip install langchain-openai
   ```

3. 涉及 Excel 操作的 COM 自动化依赖 `pywin32`（Windows）：

   ```bash
   pip install pywin32
   ```

> ⚠️ 依赖冲突提示：`requirements.txt` 中 `pandas==1.1.3`、`numpy==1.19.2` 为 2020 年旧版本，与新版 LangChain / 新 Python 在安装时可能冲突。如仅需运行本工具的核心计算与 GUI，可暂忽略 pandas/numpy；若环境安装报错，请按需升级或移除这两个旧版本。

## 配置：全局年月（config.py）

整个项目统一使用 `config.py` 中的 `YEAR` / `MONTH` 作为目标处理年月，**手动修改这里即可切换月份**，各模块不再各自用 `datetime.now()` 推算：

```python
# config.py
YEAR = 2026
MONTH = 8     # 1~12，无需补零
```

> 注意：`config` 设置的年月应与《进出场记录.xlsx》中实际数据所属月份一致，否则计算结果为空或错位。

**新增配置（区分标准人员与夜班人员）**

- **配置文件**: 在 [config.py](config.py#L1-L200) 中新增或扩展若干配置项，用以控制夜班与标准人员的加班时长与加班费计算：
   - `NIGHT_EMPLOYEE_VACATION_DAYS`（字典）：按工号映射夜班人员当月休年假天数，格式示例：{"60836": 4}（单位：天，整数）。
   - `STANDARD_EMPLOYEE_VACATION_DAYS`（可选字典）：按工号映射标准人员当月休年假天数（若需要同样的年假扣减逻辑，可在此配置）。
   - `MIN_WAGE`（数值）：最低工资（单位：元，支持小数），用于加班费基数下限，当明细表中的基本工资低于此值时使用该最低值计算时薪基数。

- **计算方法区分**:
   - **标准人员**（`calculate_dict_overtime` + `calculate_36h_truncation`）：按日累计工作日/休息日/节假日加班小时后，进入 36 小时截断分配（节假日→休息日→工作日）。如果提供 `STANDARD_EMPLOYEE_VACATION_DAYS`，可在截断前从当月累计或工作时长中扣减相应年假（实现可按需启用）。
   - **夜班人员**（`calculate_dict_overtime_night` + `calculate_night_truncation`）：夜班计算逻辑与标准人员不同（按上半夜/下半夜/白天+夜段等场景计算），且当前实现会从当月 `work_hours`（按工作日计 8 小时/天）中扣减 `NIGHT_EMPLOYEE_VACATION_DAYS` 指定的年假天数（`work_hours = max(0, work_hours - d * 15)`），再据此计算加班费与串休的分配额度。

- **加班费计算（通用）**:
   - 加班费金额仍按时薪基数 `base_rate = round(effective_salary / 21.75 / 8, 2)` 计算，`effective_salary` 为 `max(basic_salary, MIN_WAGE)`（当 `basic_salary` 无效时直接使用 `MIN_WAGE`）。此规则适用于标准与夜班人员的金额计算，保证低工资不致产生过低的基数。

- **如何设置（示例）**:

```python
NIGHT_EMPLOYEE_VACATION_DAYS = {
      "60836": 4,
}
# 若需对标准人员也生效，可添加：
STANDARD_EMPLOYEE_VACATION_DAYS = {
      "Q4642": 2,
}
MIN_WAGE = 3000.0
```

- **已修复注意事项**: 之前若误将 `NIGHT_EMPLOYEE_VACATION_DAYS` 写成集合（例如 {"60836",4}）会导致扣减逻辑被静默忽略；当前代码已使用字典格式并在读取配置时进行了保护性处理，请按示例填写。

## 使用流程

### 1. 准备源数据

将员工的进出厂打卡记录整理为 `进出场记录.xlsx`（Sheet1），包含日期时间、姓名、职号、进出厂标志等列。

### 2. 计算加班

运行主程序：

```bash
python main.py
```

弹出的窗口提供三个按钮：

- **获取月份**：弹窗输入月份（默认取自 `config.MONTH`）；
- **处理数据**：执行 `dealSheet`，清空并重置《统计表》骨架（姓名 / 职号 / 逐日日期）；
- **开始计算**：关闭窗口并触发 `calculate_main`，读取《进出场记录.xlsx》计算加班，写回《计算结果.xlsx》的 统计表 / 明细表 等。

> 推荐顺序：先点「处理数据」初始化统计表，再点「开始计算」。

### 3. 串休抵扣

```bash
python huanxiu.py
```

在窗口中：

- 输入「职号」→ 自动从《记录表》带出「姓名」；
- 「年月」默认填为 `config` 中的全局年月（格式 `YYYY-MM`，可手动修改）；
- 输入「串休时长」→ 点击「抵扣串休」，按 工作日 → 公休日 → 节假日 顺序抵扣，结果累加到统计表 K 列（负数）；
- 点击「清空 K 列」可清除指定职号+年月的 K 列抵扣记录。

### 4. AI 加班总结

`summary.py` 支持两种后端，通过 `--backend`（或 `.env` 的 `LLM_BACKEND`）切换，**默认使用本地 llama.cpp**。

**方式 A：本地 llama.cpp（默认）**

先启动本地 llama.cpp server（需支持 function calling）：

```bash
llama-server -m <你的模型路径> --port 8080
```

然后运行：

```bash
python summary.py --id Q6007
```

**方式 B：远程 DeepSeek**

在项目根目录创建 `.env` 并填写：

```ini
LLM_BACKEND=deepseek
DEEPSEEK_API_KEY=你的密钥
```

然后运行：

```bash
python summary.py --id Q6007 --backend deepseek
```

常用参数：

```bash
python summary.py --id Q6007            # 按工号总结
python summary.py --id 曲书成           # 按姓名总结
python summary.py --question "Q6007 节假日加班多少小时？"  # 自由提问
```

生成的报告写入 `reporter/report_<工号或姓名>.md`。

## 计算方法说明

本项目的加班指标由 `process_attendance.py` 计算，核心输入为《进出场记录.xlsx》的进出厂打卡记录，以及 `Crili` 提供的**日期倍率**（工作日 1.5 / 公休日 2 / 节假日 3，节假日叠加 2026 年法定节假日）。所有计算均基于 `config.py` 配置的年月。

### 1. 某天的加班时间

按人员类别分别计算当日加班小时：

- **标准人员**（`calculate_dict_overtime`）：依据当天进出厂打卡记录判定班次（白天 8:00–17:00、前半夜 17:00–24:00、后半夜 0:00–8:00 及混合班次），按各时段与排班区间的重叠秒数累加得到当日加班小时。
- **夜班人员**（`calculate_dict_overtime_night`）：按 6 种打卡场景（上半夜、下半夜、白天+前半夜、后半夜+白天、纯白班、前半夜+后半夜）计算；其中白天加班通过 `calday` 扣除午休（11:00–13:00 计 1.5 小时、其余时段计 0.5 小时），前/后半夜按 17:00–24:00、0:00–8:00 折算。

> 当日若无有效进出厂记录，或与相邻日期不闭合，则该日加班记为 0（标记为 unknown）。

### 2. 某个人某月的加班总时间

即该员工在目标年月内**每日加班时长之和**，写入《计算结果.xlsx》明细表「总加班」(E 列)，并拆分为三部分：

- 工作日加班（倍率 1.5）
- 休息日加班（倍率 2）
- 节假日加班（倍率 3）

```
总加班 = 工作日加班 + 休息日加班 + 节假日加班
```

### 3. 加班费

分为「转加班费小时」与「加班费金额」两步：

**(a) 转加班费小时（明细表 J–M 列）** —— 受 **36 小时上限** 约束（`calculate_36h_truncation`）：

- 当月转加班费的小时总数不超过 36h；
- 分配优先级：**节假日 → 休息日 → 工作日**，优先用高倍率日期填满 36h 额度；
- 未被选作加班费的部分，自动转为串休。

**(b) 加班费金额（明细表 O/P/Q/R 列）**：

- 时薪基数 `base_rate = round(基本工资 / 21.75 / 8, 2)`（基本工资取自明细表 D 列）；
- P（工作日）= `round(base_rate × 1.5 × 工作日转加班费小时)`
- Q（公休日）= `round(base_rate × 2 × 公休日转加班费小时)`
- R（节假日）= `round(base_rate × 3 × 节假日转加班费小时)`
- O（合计）= P + Q + R（金额取整）。

### 4. 可串休时间

- 「可串休时间」写入《计算结果.xlsx》记录表末列，等于该员工当月**每日转串休小时之和**；
- 转串休小时来自上面的 36h 截断规则：每天的可串休小时 = 当日加班时长 − 当日转加班费小时；当 36h 加班费额度用尽后，剩余加班时长全部转为串休；
- 「本月已扣除串休数」记录在明细表 I 列（即已通过 `huanxiu.py` 实际抵扣的小时，为负数），与「可串休时间」（应转串休总额）相互独立。

---

## 数据文件说明

### 《计算结果.xlsx》

包含以下工作表（列结构以实际文件为准）：

- **记录表**：第 1 列姓名、第 2 列职号、随后按日期排列的每日加班时长，末列为「可串休时间」；
- **统计表**：A 姓名 / B 职号 / C 日报日期 / D 考勤时间 / E 异常情况 / F 月份 / G 节假日 / H 加班或串休 / … / K 列 串休抵扣结果；
- **明细表**：每人汇总（总加班、工作日 / 休息日 / 节假日构成、转加班费、可串休、加班费金额等）；
- **汇总表**：整体汇总。

### 《进出场记录.xlsx》

源数据，Sheet1 含员工进出厂打卡明细（日期时间、姓名、职号、进/出厂标志等），由 `calculate_main` 读取。

## 注意事项

- **先备份**：`huanxiu.py`、`dealSheet`、`main.py` 都会直接读取并修改《计算结果.xlsx》，操作前请备份。
- **Windows 平台**：GUI 与 `win32com` 相关能力依赖 Windows；非 Windows 环境可能无法运行 GUI。
- **网络依赖**：`Crili` 抓取万年历、`Ccal` 获取节假日、DeepSeek 后端均需访问外网；本地 llama.cpp 不需要外网。
- **密钥安全**：`.env` 含 `DEEPSEEK_API_KEY` 等敏感信息，**请勿提交到仓库**。建议将 `.env` 加入 `.gitignore`；如需共享配置，请提供不含密钥的 `.env.example`。
- **节假日列表**：`Crili` 中 3 倍法定节假日为硬编码的 2026 年列表，跨年使用需同步更新。
- **本地模型能力**：使用 llama.cpp 后端时，所用模型需支持 function calling，否则 ReAct 工具调用无法触发。

## 许可证

当前项目未指定许可证，默认仅供内部使用。如需发布或共享，请补充合适的许可证信息。
