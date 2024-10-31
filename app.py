from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "没有文件上传", 400

    file = request.files['file']
    if file.filename == '':
        return "未选择文件", 400

    # 保存文件
    file_path = os.path.join('uploads', file.filename)
    file.save(file_path)

    return analyze_file(file_path)

def analyze_file(file_path):
    df = pd.read_excel(file_path)
    df['下单时间'] = pd.to_datetime(df['下单时间'])
    df['月份'] = df['下单时间'].dt.to_period('M')
    latest_month = df['月份'].max()

    results = {}

    for column_name in ['客户名称', '商品名称', '主营类型', '商品分类', '订单类型']:
        if column_name in df.columns:
            monthly_data = df.groupby([column_name, '月份']).agg({'实付金额': 'sum'}).reset_index()
            comparison = monthly_data[monthly_data['月份'].isin([latest_month, latest_month - 1])]

            if comparison.shape[0] >= 2:
                pivot_table = comparison.pivot(index=column_name, columns='月份', values='实付金额')
                pivot_table['环比'] = (pivot_table[latest_month] - pivot_table[latest_month - 1]) / pivot_table[latest_month - 1] * 100
                pivot_table = pivot_table.reset_index().sort_values(by='环比', ascending=False)
                results[column_name] = pivot_table

    # 保存结果到Excel
    result_file = 'analysis_result.xlsx'
    with pd.ExcelWriter(result_file) as writer:
        for key, value in results.items():
            value.to_excel(writer, sheet_name=key, index=False)

    return send_file(result_file, as_attachment=True)

if __name__ == "__main__":
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)
