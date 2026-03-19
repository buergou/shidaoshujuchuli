from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename
import io
import sys
import webbrowser
import threading
import time

# 获取程序运行路径（支持 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    # 如果是打包后的 EXE
    application_path = sys._MEIPASS
else:
    # 如果是普通 Python 脚本
    application_path = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__,
            template_folder=os.path.join(application_path, 'templates'))
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

SKIP_VALUES = {'(跳过)', '（跳过）', '跳过', '( 跳过)', '（ 跳过）'}

def process_questionnaire_data(df):
    # 固定输出17列，列名固定为指定的格式
    fixed_columns = [
        '序号', '提交答卷时间', '所用时间', '来源', '来源详情', '来自IP',
        '1、视导学科', '2、视导学校', '3、教研员', '4、目标内容',
        '5、学习方式', '6、学生表现', '7、教学效果',
        '8、作业布置（可选）', '9、科组生态', '10、是否好课堂', '11、评价与建议'
    ]

    result_rows = []

    for idx, row in df.iterrows():
        # 第7列（索引6）是"视导学科"列，前面的6列作为基本信息
        basic_info = []
        for i in range(6):  # 前6列：序号、提交答卷时间、所用时间、来源、来源详情、来自IP
            basic_info.append(row.iloc[i] if i < len(row) else '')

        # 视导学科列
        subject = row.iloc[6] if len(row) > 6 else ''
        basic_info.append(subject)

        # 视导学科之后所有列，过滤掉跳过和空值，只保留有效值
        non_skip_values = []
        for val in row.iloc[7:]:
            val_str = str(val).strip() if pd.notna(val) else ''
            if val_str not in SKIP_VALUES and val_str != '':
                non_skip_values.append(val)

        # 将有效值填入后10列，不足10个用空字符串补齐，超出部分截断
        tail = non_skip_values[:10]
        tail += [''] * (10 - len(tail))

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

def open_browser():
    """延迟打开浏览器"""
    time.sleep(2)  # 等待服务启动
    webbrowser.open('http://localhost:5000')

if __name__ == '__main__':
    # 如果是打包后的 EXE，自动打开浏览器
    if getattr(sys, 'frozen', False):
        # 在新线程中打开浏览器，避免阻塞服务器启动
        threading.Thread(target=open_browser, daemon=True).start()

    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
