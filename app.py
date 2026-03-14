from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

SKIP_VALUES = {'(跳过)', '（跳过）', '跳过', '( 跳过)', '（ 跳过）'}

def process_questionnaire_data(df):
    # 固定列名：前3列保持原样，D-M列用原始第1行（表头）的列名，固定共13列
    fixed_columns = df.columns[:13].tolist()  # A-M共13列
    num_fixed_tail = 10  # D-M共10列（索引3-12）

    result_rows = []

    for idx, row in df.iterrows():
        # 前3列（A-C）直接保留
        basic_info = row.iloc[:3].tolist()

        # D列之后所有列，过滤掉跳过和空值，只保留有效值
        non_skip_values = []
        for val in row.iloc[3:]:
            val_str = str(val).strip() if pd.notna(val) else ''
            if val_str not in SKIP_VALUES and val_str != '':
                non_skip_values.append(val)

        if non_skip_values:
            # 有效值依次填入D-M列（最多10个），超出部分截断
            tail = non_skip_values[:num_fixed_tail]
            # 不足10个用空字符串补齐
            tail += [''] * (num_fixed_tail - len(tail))
            result_rows.append(basic_info + tail)

    if not result_rows:
        return pd.DataFrame(columns=fixed_columns)

    result_df = pd.DataFrame(result_rows, columns=fixed_columns)
    result_df = result_df.dropna(how='all')

    return result_df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            df = pd.read_excel(file)
            
            original_shape = df.shape
            result_df = process_questionnaire_data(df)
            processed_shape = result_df.shape
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='处理结果')
            output.seek(0)
            
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='处理后的数据.xlsx'
            )
        except Exception as e:
            return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500
    else:
        return jsonify({'error': '请上传Excel文件（.xlsx或.xls）'}), 400

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
