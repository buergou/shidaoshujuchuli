import pandas as pd
import numpy as np

# 读取原始数据
file_path = '3526415972.xlsx'
df = pd.read_excel(file_path)

print("=" * 80)
print("数据结构分析")
print("=" * 80)

print(f"\n数据形状: {df.shape[0]}行 × {df.shape[1]}列")

print("\n前10列的列名:")
for i, col in enumerate(df.columns[:10]):
    print(f"{i}: {col}")

print("\n后10列的列名:")
for i, col in enumerate(df.columns[-10:]):
    print(f"{len(df.columns)-10+i}: {col}")

print("\n前3行数据预览:")
print(df.head(3))

print("\n检查'（跳过）'值:")
skip_count = (df == '（跳过）').sum().sum()
print(f"包含'（跳过）'的单元格数量: {skip_count}")

print("\n检查'(跳过)'值:")
skip_count2 = (df == '(跳过)').sum().sum()
print(f"包含'(跳过)'的单元格数量: {skip_count2}")

print("\nA、B、C列的前5行:")
print(df.iloc[:5, :3])

print("\n每列非空值数量（前20列）:")
print(df.iloc[:, :20].count())
