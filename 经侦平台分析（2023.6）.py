
import pandas as pd
import numpy as np
from tkinter import filedialog,messagebox
from openpyxl.workbook import Workbook
import ttkbootstrap as ttk
from PIL import ImageTk, Image
import sys
import os


# 选择数据文件夹
def open_file_a():
    global data_file
    data_file = filedialog.askdirectory()
    data_entry.insert("end", f"{data_file}")


# 选择要存放文件的位置
def open_file_b():
    global save_file_name
    save_file_name = filedialog.askdirectory()
    save_entry.insert(ttk.END, f"{save_file_name}")


# 合并csv表格
def add_csv(path_origin, use_keyword, not_use_keyword,range_columns):
    df = []
    dfs = []
    df_concat = []
    file_list = os.listdir(path_origin)
    for i in file_list:
        if use_keyword in i:
            if not not_use_keyword in i:
                df = pd.read_csv(path_origin + '\\' + i, encoding='gb18030',usecols=range(range_columns),dtype=object)
                dfs.append(df)
    df_concat = pd.concat(dfs, ignore_index=True)
    return df_concat


# 删除字符串前后的空格
def strip_whitespace(x):
    if isinstance(x, str):
        return x.strip()
    else:
        return x

# 生成资金分析表
def data_ana():
    btn_data_ana.config(state=ttk.DISABLED)
    messagebox.showinfo("安徽经侦平台资金分析", "正在处理，请稍后...")
    df = pd.read_excel(save_path + '交易流水.xlsx')
    pivot_table = pd.pivot_table(df, values='交易金额', index=['对手户名','交易对手账卡号'], columns='收付标志', aggfunc='sum')
    pivot_table = pivot_table.sort_values(by='出', ascending=False)
    pivot_table.to_excel(save_path + '交易流水资金分析.xlsx', index=True)
    messagebox.showinfo("安徽经侦平台资金分析", "已完成！")

# 主函数
def process_files():
    btn_process.config(state=ttk.DISABLED)
    # 未选择则报错
    progress_bar['maximum'] = 100
    progress_bar['value'] = 0
    folder_path = data_file

    # 定义表格名称关键字
    business_keyword = '交易明细信息'
    people_keyword = '人员信息'
    account_keyword = '账户信息'
    not_use_keyword = '子账户'
    path_source = folder_path
    path_output = save_file_name

    # 合并表格
    df_business = add_csv(path_source, business_keyword, not_use_keyword,29)
    df_account = add_csv(path_source, account_keyword, not_use_keyword,18)
    df_people = add_csv(path_source, people_keyword, not_use_keyword,14)

    # 删除前后字符串
    df_account.rename(str.strip, axis='columns', inplace=True)
    df_business.rename(str.strip, axis='columns', inplace=True)
    df_people.rename(str.strip, axis='columns', inplace=True)
    df_business = df_business.applymap(strip_whitespace)
    df_account = df_account.applymap(strip_whitespace)
    df_people = df_people.applymap(strip_whitespace)
    # 创建表格副本
    df_account_split = df_account.copy()
    # 删除某一列中字符串的 _ 之后的部分
    df_account_split['交易账号'] = df_account['交易账号'].str.split('_').str[0]
    # 删除空行
    df_business.drop("查询反馈结果原因", axis=1, inplace=True)
    df_business.replace('', np.nan, inplace=True)
    df_business = df_business.dropna(how='all')
    df_account_split = df_account_split.dropna(how='all')
    df_business_copy = df_business.copy()
    df_business_copy['交易卡号']= df_business['交易卡号'].str.split('_').str[0]
    df_business_copy['交易账号'].fillna(df_business['交易卡号'], inplace=True)

    # 更新进度条
    progress_bar['value'] += 20
    value = 20
    percentage.set(f"{int(value)}%")
    root.update()


    # 比对卡号，补全户名和证件号
    df_compare = pd.merge(left=df_business_copy, right=df_account_split, how='left', on='交易账号')
    df_compare['交易方户名'] = df_compare['账户开户名称']
    df_compare['交易方证件号码'] = df_compare['开户人证件号码']
    df_compare['交易卡号_x'].fillna(df_compare['交易卡号_y'], inplace=True)
    df_compare.insert(0, '开户银行', df_compare['账号开户银行'])
    df_compare_minus = df_compare.iloc[:, 0:29]
    df_compare_minus.rename(columns={'交易卡号_x': '交易卡号', '备注_x': '备注'}, inplace=True)

    # 比对账号，补全户名和证件号
    df_compare_minus_null = df_compare_minus[df_compare_minus['交易方户名'].isnull()]
    df_compare_minus_notnull = df_compare_minus[df_compare_minus['交易方户名'].notnull()]
    df_mid = pd.merge(left=df_compare_minus_null, right=df_account_split, how='left', on='交易卡号')
    df_mid['交易方户名'] = df_mid['账户开户名称']
    df_mid['交易方证件号码'] = df_mid['开户人证件号码']
    df_mid['开户银行'] = df_mid['账号开户银行']
    df_mid['交易账号_x'].fillna(df_mid['交易账号_y'], inplace=True)
    df_mid_minus = df_mid.iloc[:, 0:29]
    df_mid_minus.rename(columns={'交易账号_x': '交易账号', '备注_x': '备注'}, inplace=True)

    # 合并去重
    df_end = pd.concat([df_mid_minus, df_compare_minus_notnull])
    df_end_drop1 = df_end.drop_duplicates()

    # 去交易金额0.00的记录
    df_end_drop2 = df_end_drop1[~(df_end_drop1['交易金额'].isin(['0.00']))]

    # 去空行
    df_end_drop2 = df_end_drop2.dropna(subset=['交易时间'])
    # 更新进度条
    progress_bar['value'] += 20
    value += 20
    percentage.set(f"{int(value)}%")
    root.update()
    # 去负号
    df_end_drop3 =df_end_drop2.copy()
    df_end_drop3['交易金额'] = df_end_drop2['交易金额'].map(lambda x: abs(float(x)))
    data_name  = os.path.basename(data_file)
    global save_path
    save_path=os.path.join(path_output, data_name)
    # 更新进度条
    progress_bar['value'] += 20
    value += 20
    percentage.set(f"{int(value)}%")
    root.update()
    df_end_drop3.to_excel(save_path + '交易流水.xlsx', index=False, float_format="%.2f")
    # 更新进度条
    progress_bar['value'] += 20
    value += 20
    percentage.set(f"{int(value)}%")
    root.update()
    df_account_split.to_excel(save_path + '账户信息.xlsx', index=False, float_format='@')
    df_people.to_excel(save_path + '人员信息.xlsx', index=False, float_format='@')
    # 更新进度条
    progress_bar['value'] += 20
    value += 20
    percentage.set(f"{int(value)}%")
    root.update()
    messagebox.showinfo("安徽经侦平台资金分析", "已完成！")
    btn_data_ana.config(state=ttk.NORMAL)


if __name__ == '__main__':
    # 选择需要清洗的文件夹
    root = ttk.Window()
    root.title("安徽经侦平台资金分析")
    root.geometry("1000x800")
    # 拼接图片文件的完整路径
    current_path = os.path.dirname(sys.argv[0])
    image_path = os.path.join(current_path, '1.png')
    image = Image.open(image_path)  # 替换为你的图片路径
    image = image.resize((1000, 200))  # 调整图片大小
    photo = ImageTk.PhotoImage(image)
    # 创建一个Label组件并设置图片属性
    img_label = ttk.Label(root, image=photo)
    img_label.place(x=0,y=0)
    # 创建选择总表的label、button和entry
    data_label = ttk.Label(root, text="请选择表格所在文件夹:")
    data_label.place(x=0,y=300)
    data_entry = ttk.Entry(root,width=40, bootstyle="info")
    data_entry.place(x=300,y=300)
    data_btn = ttk.Button(root, text="选择", command=open_file_a, bootstyle="info")
    data_btn.place(x=880,y=300)
    # 创建存放文件的label、button和entry
    save_label = ttk.Label(root, text="请选择要存放的文件夹:")
    save_label.place(x=0,y=400)
    save_entry = ttk.Entry(root,width=40,bootstyle="info")
    save_entry.place(x=300,y=400)
    save_btn = ttk.Button(root, text="选择", command=open_file_b)
    save_btn.place(x=880,y=400)
    # 创建处理文件的按钮
    btn_process = ttk.Button(root, text="生成excel表格", command=process_files, bootstyle="success")
    btn_process.place(x=300,y=500)
    data_ana_label = ttk.Label(root, text="请在生成excel表格后选择：")
    data_ana_label.place(x=0, y=700)
    btn_data_ana = ttk.Button(root, text="生成交易对方资金分析表", command=data_ana, bootstyle="success")
    btn_data_ana.place(x=300, y=700)
    btn_data_ana.config(state=ttk.DISABLED)
    # 进度条
    progress_label = ttk.Label(root, text="当前进度：")
    progress_label.place(x=0,y=600)
    progress_bar = ttk.Progressbar(root, length=600, mode='determinate',bootstyle="success-striped")
    progress_bar.place(x=300,y=600)
    percentage = ttk.StringVar()
    percentage_label = ttk.Label(root, textvariable=percentage)
    percentage_label.place(x=880,y=600)
    root.mainloop()
