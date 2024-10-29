import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO

# 设置页面标题和 LOGO
st.set_page_config(page_title="大麦-数据策略工具-购买周期", layout="centered")

# 加载并显示 LOGO 图片
logo = Image.open("logo.png").resize((200, 70))
st.image(logo)

st.title("大麦-数据策略工具-购买周期")

# 上传 Excel 文件
uploaded_file = st.file_uploader("请上传需要分析的表格文件：", type=["xlsx"])

# 输入商品名称
product_name = st.text_input("请输入要查询的商品名称：")

if uploaded_file and product_name:
    try:
        # 加载 Excel 数据
        data = pd.ExcelFile(uploaded_file)
        df = data.parse(data.sheet_names[0])

        # 严格匹配商品名称
        filtered_data = df[df['商品名称'] == product_name]
        if filtered_data.empty:
            st.warning("查询不到此商品，请重新输入。")
        else:
            # 转换时间格式并排序
            filtered_data['下单时间'] = pd.to_datetime(filtered_data['下单时间'], errors='coerce')
            filtered_data = filtered_data.sort_values(['客户名称', '下单时间'])

            # 去除无效时间数据
            filtered_data = filtered_data.dropna(subset=['下单时间'])

            # 计算购买周期
            filtered_data['购买间隔(天)'] = filtered_data.groupby('客户名称')['下单时间'].diff().dt.days

            # 最近一次下单时间
            recent_order = filtered_data.groupby('客户名称')['下单时间'].max().reset_index()
            recent_order.rename(columns={'下单时间': '最近一次下单时间'}, inplace=True)
            recent_order['最近一次下单时间'] = recent_order['最近一次下单时间'].dt.strftime('%m月%d日')

            # 汇总数据
            summary = filtered_data.groupby(['客户名称', 'BD'])['购买间隔(天)'].agg(['mean', 'min', 'max']).reset_index()
            summary.rename(columns={'mean': '平均购买周期(天)', 'min': '最短购买周期(天)', 'max': '最长购买周期(天)'}, inplace=True)

            # 合并最近一次下单时间
            summary = pd.merge(summary, recent_order, on='客户名称', how='left')

            # 预测购买时间
            summary['预测购买时间'] = pd.to_datetime(summary['最近一次下单时间'], format='%m月%d日', errors='coerce') + pd.to_timedelta(summary['平均购买周期(天)'], unit='D')
            summary['预测购买时间'] = summary['预测购买时间'].dt.strftime('%m月%d日').fillna('无')

            summary['商品名称'] = product_name

            # **过滤掉所有购买周期都为 0 或 NaN 的客户**
            summary = summary[
                (summary['平均购买周期(天)'].fillna(0) != 0) |
                (summary['最短购买周期(天)'].fillna(0) != 0) |
                (summary['最长购买周期(天)'].fillna(0) != 0)
            ]

            # 显示结果表
            st.dataframe(summary)

            # 提供 Excel 下载
            output = BytesIO()
            summary.to_excel(output, index=False, engine='xlsxwriter')
            output.seek(0)

            st.download_button(
                label="下载结果表格",
                data=output,
                file_name="购买周期分析结果.xlsx",
                mime="application/vnd.ms-excel"
            )
    except Exception as e:
        st.error(f"出现错误：{e}")
