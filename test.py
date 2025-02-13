import pandas as pd

# 构造示例数据
data = {
'Name': ['Alice', 'Bob', 'Charlie'],
'Age': [25, 30, 35],
'City': ['New York', 'Los Angeles', 'Chicago']
}
df = pd.DataFrame(data)

# 使用 iterrows() 逐行迭代
for index, row in df.iterrows():
    print(f"Index: {index}")
    print(f"Row data:\n{row}\n")