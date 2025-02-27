import numpy as np
import datetime
n = {'20250122': 1.45, '20250118': 7.47, '20250117': 1.6, '20250116': 0.57, '20250114': 2.0, '20250113': 0.78, '20250112': 7.78, '20250109': 2.78, '20250108': 2.17, '20250107': 2.3, '20250106': 2.22, '20250104': 8.4, '20250103': 1.97, '20250102': 2.27}

# 定义节假日日期（需要根据实际情况补充）
holidays = ['20250101', '20250102', '20250103']  # 示例节假日

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
    result.append([date_str, modified_value,float(value)])

result_array = np.array(result)


def maximize_value(result_array, max_sum=36):
    # 将字符串转换为数值
    values = result_array[:, 2].astype(float)
    coefficients = result_array[:, 1].astype(float)
    
    n = len(values)
    dp = np.zeros((n+1, max_sum+1))
    
    # 填充DP表
    for i in range(1, n+1):
        for j in range(max_sum+1):
            if values[i-1] <= j:
                dp[i][j] = max(dp[i-1][j], 
                             dp[i-1][j-int(values[i-1])] + coefficients[i-1]*values[i-1])
            else:
                dp[i][j] = dp[i-1][j]
    
    # 回溯找到选择的数组
    j = max_sum
    selected = []
    total = 0
    for i in range(n, 0, -1):
        if dp[i][j] != dp[i-1][j] and (total + values[i-1]) <= max_sum:
            selected.append(result_array[i-1])
            total += values[i-1]
            j -= int(values[i-1])
    
    # 确保总和不超过36
    selected_array = np.array(selected)
    assert sum(selected_array[:, 2].astype(float)) <= max_sum, "总和超过36"
    
    return np.array(selected)


print(sum(maximize_value(result_array)[:,2].astype(float)))
