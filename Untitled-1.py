def maximize_value(result, max_sum=36000):
    # 提取重量和价值
    weights = [int(row[2]*1000) for row in result]
    values = [int(row[1] * row[2]*1000) for row in result]
    n = len(weights)

    # 创建动态规划表
    dp = [[0] * (max_sum + 1) for _ in range(n + 1)]

    # 填充动态规划表
    for i in range(1, n + 1):
        for w in range(max_sum + 1):
            if weights[i-1] <= w:
                dp[i][w] = max(dp[i-1][w], dp[i-1]
                               [w - weights[i-1]] + values[i-1])
            else:
                dp[i][w] = dp[i-1][w]

    # 回溯找到选择的物品
    w = max_sum
    selected_items = []
    for i in range(n, 0, -1):
        if dp[i][w] != dp[i-1][w]:
            selected_items.append(result[i-1])
            w -= weights[i-1]

    return selected_items, dp[n][max_sum]/1000


# 测试数据
result = [
    ['20250122', 1.5, 1.45], ['20250118', 2, 7.47], ['20250117', 1.5, 1.6],
    ['20250116', 1.5, 0.57], ['20250114', 1.5, 2.0], ['20250113', 1.5, 0.78],
    ['20250112', 2, 7.78], ['20250109', 1.5, 2.78], ['20250108', 1.5, 2.17],
    ['20250107', 1.5, 2.3], ['20250106', 1.5, 2.22], ['20250104', 2, 8.4],
    ['20250103', 1.5, 1.97], ['20250102', 1.5, 2.27]
]

# 输出结果
print('Maximized value:', maximize_value(result))
