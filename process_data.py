import pandas as pd
import numpy as np

def process_questionnaire_data(input_file, output_file):
    """
    处理问卷数据，将分散的学科题目整合到固定列位置
    
    参数:
        input_file: 输入文件路径
        output_file: 输出文件路径
    """
    
    # 读取原始数据
    print(f"正在读取文件: {input_file}")
    df = pd.read_excel(input_file)
    
    print(f"原始数据形状: {df.shape[0]}行 × {df.shape[1]}列")
    
    # 创建结果DataFrame，保留A、B、C列
    result_data = []
    
    # 遍历每一行
    for idx, row in df.iterrows():
        # 保留前三列（A、B、C列）
        basic_info = row.iloc[:3].tolist()
        
        # 找到该行非"(跳过)"的值（从第4列开始）
        non_skip_values = []
        for val in row.iloc[3:]:
            # 检查是否为跳过值（考虑多种格式）
            val_str = str(val).strip() if pd.notna(val) else ''
            if val_str not in ['(跳过)', '（跳过）', '跳过', '( 跳过)', '（ 跳过）', '']:
                non_skip_values.append(val)
        
        # 如果找到非跳过值，添加到结果中
        if len(non_skip_values) > 0:
            new_row = basic_info + non_skip_values
            result_data.append(new_row)
    
    # 创建结果DataFrame
    # 列名：前三列使用原始列名，后面的列根据实际数量命名
    max_questions = max(len(row) - 3 for row in result_data) if result_data else 0
    columns = df.columns[:3].tolist() + [f'题目{i+1}' for i in range(max_questions)]
    
    # 确保所有行的长度一致
    for row in result_data:
        while len(row) < len(columns):
            row.append('')
    
    result_df = pd.DataFrame(result_data, columns=columns)
    
    # 删除完全空白的行
    result_df = result_df.dropna(how='all')
    
    print(f"处理后数据形状: {result_df.shape[0]}行 × {result_df.shape[1]}列")
    
    # 保存到新文件
    result_df.to_excel(output_file, index=False)
    print(f"已保存到: {output_file}")
    
    return result_df

if __name__ == "__main__":
    # 输入和输出文件路径
    input_file = '3526415972.xlsx'
    output_file = 'xin.xlsx'
    
    # 处理数据
    result = process_questionnaire_data(input_file, output_file)
    
    # 显示处理结果预览
    print("\n处理结果预览（前5行）:")
    print(result.head())
    
    print("\n处理结果预览（后5行）:")
    print(result.tail())
