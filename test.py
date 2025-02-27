import numpy as np
import datetime
import itertools
import time

start = time.time()
n = {'20250122': 1.45, '20250118': 7.47, '20250117': 1.6, '20250116': 0.57, '20250114': 2.0, '20250113': 0.78, '20250112': 7.78,
     '20250109': 2.78, '20250108': 2.17, '20250107': 2.3, '20250106': 2.22, '20250104': 8.4, '20250103': 1.97, '20250102': 2.27}
# 定义节假日日期（需要根据实际情况补充）
holidays = ['20250101', '20250128', '20250129',
            '20250130', '20250131']  # 示例节假日

# 创建numpy数组
result = []
for date_str, value in n.items():
    date_obj = datetime.datetime.strptime(date_str, '%Y%m%d')
    if date_str in holidays:
        modified_value = 3
    elif date_obj.weekday() >= 5:  # 周末
        modified_value = 2
    else:  # 工作日
        modified_value = 1.5
    result.append([date_str, modified_value, float(value)])


def maximize_value(result, max_sum=36):
    # 将字符串转换为数值
    max = 0
    for i in range(2, len(result)+1):
        combinations = list(itertools.combinations(result, i))
        for combination in combinations:
            n = np.array(combination)
            if sum(n[:, 2].astype(float)) <= max_sum:
                s = sum(n[:, 1].astype(float)*n[:, 2].astype(float))
                if s > max:
                    max = s
                    x_max = n
    print(x_max, max)

    return max


maximize_value(result)
print('Time:', time.time()-start)
