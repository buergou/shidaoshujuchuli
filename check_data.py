import pandas as pd

# 读取处理后的文件
df = pd.read_excel('处理后的数据 (12).xlsx')

print(f'行数: {len(df)}')
print(f'列数: {len(df.columns)}')
print('\n列名:')
for i, col in enumerate(df.columns):
    print(f'  {i}: {col}')

print(f'\n前5行数据:')
print(df.head().to_string())

print(f'\n\n数据类型:')
print(df.dtypes)
