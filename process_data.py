import pandas as pd

# 跳过值集合
SKIP_VALUES = {'(跳过)', '（跳过）', '跳过', '( 跳过)', '（ 跳过）'}

def process_questionnaire_data(df):
    # 打印原始文件的列名,用于调试
    print("\n" + "=" * 60)
    print("【调试】原始文件的列名:")
    for i, col in enumerate(df.columns):
        print(f"  列{i}: {col}")
    print("=" * 60 + "\n")

    # 固定输出17列,列名固定为指定的格式
    fixed_columns = [
        '序号', '提交答卷时间', '所用时间', '来源', '来源详情', '来自IP',
        '1、视导学科', '2、视导学校', '3、教研员', '4、目标内容',
        '5、学习方式', '6、学生表现', '7、教学效果',
        '8、作业布置(可选)', '9、科组生态', '10、是否好课堂', '11、评价与建议'
    ]

    result_rows = []

    for idx, row in df.iterrows():
        # 第7列(索引6)是"视导学科"列,前面的6列作为基本信息
        basic_info = []
        for i in range(6):  # 前6列:序号、提交答卷时间、所用时间、来源、来源详情、来自IP
            basic_info.append(row.iloc[i] if i < len(row) else '')

        # 视导学科列
        subject = row.iloc[6] if len(row) > 6 else ''
        basic_info.append(subject)

        # 视导学科之后所有列,过滤掉跳过和空值,只保留有效值
        non_skip_values = []
        for val in row.iloc[7:]:
            val_str = str(val).strip() if pd.notna(val) else ''
            if val_str not in SKIP_VALUES and val_str != '':
                non_skip_values.append(val)

        # 将有效值填入后10列,不足10个用空字符串补齐,超出部分截断
        tail = non_skip_values[:10]
        tail += [''] * (10 - len(tail))

        result_rows.append(basic_info + tail)

    if not result_rows:
        return pd.DataFrame(columns=fixed_columns)

    result_df = pd.DataFrame(result_rows, columns=fixed_columns)
    result_df = result_df.dropna(how='all')

    # 【测试标记】检查代码是否执行
    print("\n" + "=" * 60)
    print("【重要】数据处理函数已执行！！！")
    print("=" * 60 + "\n")

    # 按"2、视导学校"列排序,相同学校的数据集中在一起
    result_df['2、视导学校'] = result_df['2、视导学校'].astype(str)

    print(f"【调试】排序前的学校列(前10个): {result_df['2、视导学校'].head(10).tolist()}")
    print(f"【调试】排序前的学校列(后10个): {result_df['2、视导学校'].tail(10).tolist()}")

    result_df = result_df.sort_values(by='2、视导学校', na_position='last')
    result_df = result_df.reset_index(drop=True)

    print(f"【调试】排序后的学校列(前10个): {result_df['2、视导学校'].head(10).tolist()}")
    print(f"【调试】排序后的学校列(后10个): {result_df['2、视导学校'].tail(10).tolist()}")
    print(f"【调试】总行数: {len(result_df)}")

    # 重新生成序号,保持连续性
    result_df['序号'] = range(1, len(result_df) + 1)

    return result_df

if __name__ == '__main__':
    # 读取原始Excel文件
    input_file = '353230132_按文本_2026年春季学期小学教研部听课视导反馈_14_14.xlsx'
    output_file = '处理后的数据_已排序.xlsx'

    print(f"正在读取文件: {input_file}")
    df = pd.read_excel(input_file)
    print(f"原始数据: {df.shape[0]} 行 x {df.shape[1]} 列")

    # 处理数据
    result_df = process_questionnaire_data(df)

    # 保存结果
    result_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\n处理完成!结果已保存到: {output_file}")
    print(f"输出数据: {result_df.shape[0]} 行 x {result_df.shape[1]} 列")

    # 显示每所学校的数据量统计
    school_counts = result_df['2、视导学校'].value_counts()
    print(f"\n各学校数据统计:")
    for school, count in school_counts.items():
        if school and school != 'nan':
            print(f"  {school}: {count} 条")
