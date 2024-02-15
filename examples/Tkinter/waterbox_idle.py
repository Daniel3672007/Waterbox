#!/usr/bin/python3
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from datetime import datetime
import influxdb_client
from influxdb_client import InfluxDBClient
from influxdb_client.client.write_api import SYNCHRONOUS
import pandas as pd
import os

# 连接InfluxDB
token = "fP-GBq8Z1wZE7iW8qFBuxVy-ArVP9TqVec0naJ77XLECiwSr82aRXqvo3ylXZqU_2ad2vxWGcMoMbl3PXqAZ7A=="
server_url = "http://140.112.12.62:8086"
client = InfluxDBClient(url=server_url, token=token)
query_api = client.query_api()

# 查询参数
params = ["bat_a", "bat_v", "s_ec2", "s_ph2", "date", "s_Tb", "s_t7"]  # 参数列表

# 创建Tkinter窗口
root = tk.Tk()
root.title("Matplotlib in Tkinter")

# 创建Matplotlib图形
fig = Figure(figsize=(8, 4), dpi=100)
ax = fig.add_subplot(111)

# 创建FigureCanvasTkAgg实例
canvas = FigureCanvasTkAgg(fig, master=root)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack()

# 创建下拉菜单
param_var = tk.StringVar(value=params[0])
param_menu = ttk.Combobox(root, textvariable=param_var, values=params)
param_menu.pack(pady=10)

global result

# 查询并绘图函数
def query_and_plot():
    global result
    try:
        param = param_var.get()
        query = f'from(bucket: "WaterBox")\
                |> range(start: -30d)\
                |> filter(fn:(r) => r.device_id == "9C65F9448BD3")\
                |> filter(fn:(r) => r._field == "{param}")\
                |> drop(columns: ["_start", "_stop"])'

        result = query_api.query(org="NTUCE", query=query)

        # 清除現有圖表
        ax.clear()

        for table in result:
            field_name = table.records[0].values['_field']
            data = [data.values['_value'] for data in table.records]
            time = [data.values['_time'] for data in table.records]
            ax.set_title(field_name)
            ax.set_xlabel("Time")
            ax.set_ylabel("Value")
            ax.grid(True)
            ax.plot(time, data, label=field_name)

        # 显示图表
        canvas.draw()

        # 調整子圖位置
        fig.tight_layout()

    except tk.TclError:
        pass  # 忽略 TclError


# 导出至Excel按钮
def export_to_excel():
    if 'result' in globals() and result:
        # 創建空列表以存儲所有資料
        data_list = []

        # 逐一遍歷查詢結果並提取資料和時間戳記
        for table in result:
            field_name = table.records[0].values['_field']
            data = [data.values['_value'] for data in table.records]
            time = [data.values['_time'] for data in table.records]
            # 將資料存入data_list列表
            data_list.extend({'Field Name': field_name, 'Time': t, 'Data': d} for t, d in zip(time, data))
        
        # 將data_list轉換為pandas的DataFrame
        df = pd.DataFrame(data_list)

        # 使用pivot_table進行資料重組
        pivot_df = df.pivot_table(index='Time', columns='Field Name', values='Data', aggfunc='first')

        # Calculate the 'omega' column by dividing 'bat_a' by 'bat_v'
        if 'bat_a' in pivot_df.columns and 'bat_v' in pivot_df.columns:
            pivot_df['omega'] = pivot_df['bat_a'] / pivot_df['bat_v']

        # 將時間戳記轉換為字串
        pivot_df.index = pivot_df.index.astype(str)

        # 將DataFrame寫入Excel檔案
        excel_filename = 'data_output_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'
        pivot_df.to_excel(excel_filename, index=True)
        messagebox.showinfo("Export to Excel", f"資料已成功寫入Excel檔案: {excel_filename}")
        print("資料已成功寫入Excel檔案:", excel_filename)


# ------------------------------------------------------------------- #
# 创建查询按钮
query_button = tk.Button(root, text="繪製趨勢圖", command=query_and_plot)
query_button.pack()

export_button = tk.Button(root, text="導出至Excel", command=export_to_excel)
export_button.pack()

# 启动Tkinter主循环
root.mainloop()
