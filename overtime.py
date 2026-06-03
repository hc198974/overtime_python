from demos import *
import sys
from openpyxl import load_workbook
from overtime_criterion import Count_criterion
from overtime_night import Count_night

if __name__ == "__main__":
    cw = Cwindow()
    if not cw.createWindow():
        print("已取消计算：窗口已关闭")
        sys.exit(0)

    night_num = ['Q4642', '60836']  # 60836
    # 获得工作日和节假日
    result = Crili(2026, cw.month).parseHTML()
    start = time.perf_counter()
    wb = load_workbook(filename="计算结果.xlsx")
    ws = wb["记录表"]
    # 用id取代name，识别职号
    ids_criterion = []
    ids_night = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=2):
        for id in row:
            if id.value is not None and id.value not in night_num:
                ids_criterion.append(id)
            elif id.value in night_num:
                ids_night.append(id)

    Count_criterion(ids_criterion, cw.month, result, wb).jiSuan()
    Count_night(ids_night, cw.month, result, wb).jiSuan()

    wb.save("计算结果.xlsx")

    print("运行时间：", time.perf_counter() - start)
