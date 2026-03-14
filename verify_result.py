import pandas as pd

# 读取处理后的数据
file_path = 'xin.xlsx'
df = pd.read_excel(file_path)

print("=" * 80)
print("处理结果验证")
print("=" * 80)

print(f"\n数据形状: {df.shape[0]}行 × {df.shape[1]}列")

print("\n列名:")
for i, col in enumerate(df.columns):
    print(f"{i}: {col}")

print("\n完整数据预览:")
print(df.to_string())

print("\n检查是否还有'(跳过)'值:")
skip_count = (df == '(跳过)').sum().sum()
print(f"包含'(跳过)'的单元格数量: {skip_count}")

print("\n检查是否还有'（跳过）'值:")
skip_count2 = (df == '（跳过）').sum().sum()
print(f"包含'（跳过）'的单元格数量: {skip_count2}")

print("\n每列非空值数量:")
print(df.count())
