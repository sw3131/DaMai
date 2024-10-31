import streamlit as st
import pandas as pd

# 设置页面标题
st.title("大麦-数据与策略-月环比智能")

# 文件上传
uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx"])

if uploaded_file is not None:
    # 读取Excel文件
    df = pd.read_excel(uploaded_file)
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

    # 显示分析结果
    for key, value in results.items():
        st.subheader(key)
        st.dataframe(value)

    # 下载按钮
    if st.button("下载分析结果"):
        result_file = "analysis_result.xlsx"
        with pd.ExcelWriter(result_file) as writer:
            for key, value in results.items():
                value.to_excel(writer, sheet_name=key, index=False)
        with open(result_file, "rb") as file:
            st.download_button("下载", file, result_file)
