import pandas as pd
import itertools

# Create a sample DataFrame
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35]
}
df = pd.DataFrame(data)

# Generate all possible combinations of rows taken 2 at a time
combinations = list(itertools.combinations(df.itertuples(), 3))

# Print the combinations
for combination in combinations:
    print(combination)

