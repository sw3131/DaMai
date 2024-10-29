import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from PIL import Image, ImageTk  # 用于加载LOGO图片

# 全局变量，用于保存上传的文件路径
uploaded_file_path = None

def analyze_data(file_path, product_name):
    data = pd.ExcelFile(file_path)
    sheet_name = data.sheet_names[0]
    df = data.parse(sheet_name)

    # 根据商品名称严格匹配过滤数据
    filtered_data = df[df['商品名称'] == product_name]
    if filtered_data.empty:
        messagebox.showwarning("提示", "查询不到此商品，请重新输入。")
        return None

    # 转换下单时间格式并排序
    filtered_data['下单时间'] = pd.to_datetime(filtered_data['下单时间'])
    filtered_data = filtered_data.sort_values(['客户名称', '下单时间'])

    # 计算客户购买周期
    filtered_data['购买间隔(天)'] = filtered_data.groupby('客户名称')['下单时间'].diff().dt.days

    # 获取每位客户的最近一次下单时间
    recent_order = filtered_data.groupby('客户名称')['下单时间'].max().reset_index()
    recent_order.rename(columns={'下单时间': '最近一次下单时间'}, inplace=True)
    recent_order['最近一次下单时间'] = recent_order['最近一次下单时间'].dt.strftime('%m月%d日')

    # 汇总结果并增加 BD 信息和商品名称
    summary = filtered_data.groupby(['客户名称', 'BD'])['购买间隔(天)'].agg(['mean', 'min', 'max']).reset_index()
    summary.rename(columns={'mean': '平均购买周期(天)', 'min': '最短购买周期(天)', 'max': '最长购买周期(天)'}, inplace=True)

    summary = pd.merge(summary, recent_order, on='客户名称', how='left')
    summary['预测购买时间'] = pd.to_datetime(summary['最近一次下单时间'], format='%m月%d日') + pd.to_timedelta(summary['平均购买周期(天)'], unit='D')
    summary['预测购买时间'] = summary['预测购买时间'].dt.strftime('%m月%d日')

    summary['商品名称'] = product_name

    summary = summary.dropna(subset=['平均购买周期(天)', '最短购买周期(天)', '最长购买周期(天)'])
    summary = summary[(summary['平均购买周期(天)'] != 0) | (summary['最短购买周期(天)'] != 0) | (summary['最长购买周期(天)'] != 0)]

    return summary

def save_file(data):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="保存结果表")
    if save_path:
        data.to_excel(save_path, index=False)
        messagebox.showinfo("保存成功", f"结果已保存至：{save_path}")

def upload_file():
    global uploaded_file_path
    uploaded_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="上传表格文件")
    if uploaded_file_path:
        messagebox.showinfo("上传成功", "文件上传成功，请输入商品名称并点击查询。")
    else:
        messagebox.showwarning("上传失败", "未选择任何文件，请重新上传。")

def query_product():
    if not uploaded_file_path:
        messagebox.showwarning("警告", "请先上传文件。")
        return

    product_name = entry.get().strip()
    if not product_name:
        messagebox.showwarning("警告", "请输入要查询的商品名称。")
        return

    result = analyze_data(uploaded_file_path, product_name)
    if result is not None:
        display_data(result)

def display_data(data):
    result_window = tk.Toplevel(root)
    result_window.title("分析结果")

    from pandastable import Table
    frame = tk.Frame(result_window)
    frame.pack(fill='both', expand=True)
    pt = Table(frame, dataframe=data)
    pt.show()

    save_btn = tk.Button(result_window, text="保存结果", command=lambda: save_file(data))
    save_btn.pack(pady=10)

# 初始化GUI窗口
root = tk.Tk()
root.title("大麦-数据策略工具-购买周期")

# 设置窗口大小和居中显示
window_width, window_height = 800, 600  # 窗口调整为800x600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 加载并显示LOGO图片
logo = Image.open("logo.png")
logo = logo.resize((200, 70), Image.LANCZOS)  # LOGO调整为宽200，高70
logo_img = ImageTk.PhotoImage(logo)
logo_label = tk.Label(root, image=logo_img)
logo_label.pack(pady=10)

# 上传文件按钮和标签
upload_label = tk.Label(root, text="上传需要分析的表格文件：", font=("Arial", 14))
upload_label.pack(pady=5)

upload_btn = tk.Button(root, text="上传文件", command=upload_file, width=15, font=("Arial", 12))
upload_btn.pack(pady=5)

# 输入商品名称标签和输入框
entry_label = tk.Label(root, text="请输入要查询的商品名称：", font=("Arial", 14))
entry_label.pack(pady=5)

entry = tk.Entry(root, width=30, font=("Arial", 12))
entry.pack(pady=5)

# 查询按钮
query_btn = tk.Button(root, text="查询", command=query_product, width=15, font=("Arial", 12))
query_btn.pack(pady=10)

# 启动主事件循环
root.mainloop()
