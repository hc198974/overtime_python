import pandas as pd

# 创建示例 DataFrame
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Los Angeles', 'Chicago']
}
df = pd.DataFrame(data)

# 将默认索引和 'Age' 列转换为字典
result_dict = df['Age'].to_dict()
print(result_dict)
