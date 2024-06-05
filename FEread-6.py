import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import pandas as pd

# 定义全局变量保存文件路径
file_path = ""
save_path = ""
original_data = {}  # 用于保存读取的Excel数据

def open_file():
    global file_path, selected_file_label, original_data
    # 打开文件对话框
    file_path = filedialog.askopenfilename(filetypes=[("DAT files", "*.dat")])

    # 如果用户取消选择文件，则直接返回
    if not file_path:
        text_output.insert(tk.END, "未选择文件\n")
        return

    # 更新选择的文件名标签
    selected_file_label.config(text="选择的文件：" + file_path)
    text_output.insert(tk.END, "文件已加载\n")

def save_excel():
    global file_path, save_path, original_data
    if not file_path:
        text_output.insert(tk.END, "请先选择要打开的文件\n")
        return

    # 读取数据文件
    with open(file_path, "r") as f:
        lines = f.readlines()

    # 找到表格的起始位置
    table_start_indices = [i for i, line in enumerate(lines) if line.strip().startswith("Table")]

    # 获取保存路径
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        text_output.insert(tk.END, "未选择保存路径\n")
        return

    # 更新保存的文件名标签
    saved_file_label.config(text="保存的文件：" + save_path)

    # 创建一个 Pandas ExcelWriter 对象
    excel_writer = pd.ExcelWriter(save_path, engine='xlsxwriter')

    # 创建一个新的 Pandas ExcelWriter 对象用于保存头部信息
    header_save_path = save_path[:-5] + "_v2.xlsx"
    header_excel_writer = pd.ExcelWriter(header_save_path, engine='xlsxwriter')

    # 创建一个新的 Pandas ExcelWriter 对象用于保存新整理的数据
    new_save_path = save_path[:-5] + "_v1.xlsx"
    new_excel_writer = pd.ExcelWriter(new_save_path, engine='xlsxwriter')

    # 获取总表格数量
    total_tables = len(table_start_indices)
    progress_step = 100 / total_tables

    # 更新进度条的函数
    def update_progress(progress):
        progress_value.set(progress)
        progress_label.config(text=f"{progress:.1f}%")  # 更新进度标签显示百分比
        root.update_idletasks()  # 强制更新界面

    # 更新文本框的函数
    def update_text(text):
        text_output.insert(tk.END, text)
        text_output.see(tk.END)  # 自动滚动到末尾
        root.after(100, scroll_to_end)  # 每100毫秒调用一次 scroll_to_end 函数

    def scroll_to_end():
        text_output.yview(tk.END)  # 将文本框滚动到最下面

    # 遍历每个表格
    for i, start_index in enumerate(table_start_indices):
        # 提取工作表名称，以 "Table" 开始后面的数字
        table_number = i + 1
        table_name = f"Table_{table_number // 2 + 1}"

        # 读取表格元数据
        metadata = {}
        table_data_start = start_index + 1
        for line in lines[start_index + 1:]:
            if line.strip() == "":
                table_data_start += 1
                continue
            if line.strip().startswith("Time [s]"):
                break
            if ":" in line:
                key, value = line.strip().split(":", 1)
                metadata[key.strip()] = value.strip()  # 保留键的原始大小写

            table_data_start += 1

        # 获取 Thickness [nm] 的值
        thickness_nm = float(metadata.get("Thickness [nm]", 1e-4))  # 如果找不到 Thickness [nm]，默认值为1e-4

        # 将表格元数据转换为 DataFrame 并保存到新的工作表中
        metadata_df = pd.DataFrame(list(metadata.items()), columns=['Key', 'Value'])
        metadata_df.to_excel(header_excel_writer, sheet_name=table_name, index=False)

        # 读取表格数据
        table_data = []
        for line in lines[table_data_start:]:
            if line.strip().startswith("Table"):
                break
            if line.strip() == "":
                continue
            data_row = line.strip().split("\t")
            if len(data_row) > 1:  # 确保行中有多个数据
                table_data.append(data_row)

        # 更新文本框
        text_output.insert(tk.END, f"表格 {table_name} 数据行数：{len(table_data)}\n")

        # 更新进度条
        current_progress = (i + 1) * progress_step
        update_progress(current_progress)

        # 如果没有有效数据，则跳过当前表格
        if len(table_data) <= 1:
            text_output.insert(tk.END, f"表格 {table_name} 中没有有效数据，跳过保存\n")
            continue

        # 将表格数据转换为DataFrame
        df = pd.DataFrame(table_data[1:], columns=table_data[0])

        # 将数字内容转换为数字类型
        df = df.apply(pd.to_numeric, errors='coerce')  # 使用'coerce'处理错误，将无法转换为数字的值设置为NaN

        # 检查是否存在 'V+ [V]' 列
        if 'V+ [V]' not in df.columns:
            text_output.insert(tk.END, f"表格 {table_name} 中缺少 'V+ [V]' 列，跳过保存\n")
            continue

        # 计算新列 E[kV]
        df['E [kV/cm]'] = df['V+ [V]'].astype(float) / thickness_nm * 1e4

        # 计算新列 Strain1 [%], Strain2 [%], Strain3 [%]
        df['Strain1 [%]'] = df['D1 [nm]'].astype(float) / thickness_nm * 100
        df['Strain2 [%]'] = df['D2 [nm]'].astype(float) / thickness_nm * 100
        df['Strain3 [%]'] = df['D3 [nm]'].astype(float) / thickness_nm * 100

        # 获取 Temperature [ C] 的值
        temperature_c = metadata.get("Temperature [ C]", "none")

        # 将 Temperature [ C] 值添加到表格名称后面
        table_name_with_temperature = f"{table_name}_T_{temperature_c}"

        # 将 DataFrame 写入 Excel 文件中的不同表格
        df.to_excel(excel_writer, sheet_name=table_name, index=False)

        # 重新按指定顺序排列列
        df_new = df[['E [kV/cm]', 'P1 [uC/cm2]', 'P2 [uC/cm2]', 'P3 [uC/cm2]', 'I1 [A]', 'I2 [A]', 'I3 [A]', 'Strain1 [%]', 'Strain2 [%]', 'Strain3 [%]']]
        
        # 保存新整理的数据到新的 Excel 文件中
        df_new.to_excel(new_excel_writer, sheet_name=table_name_with_temperature, index=False)

        # 更新文本框
        text_output.insert(tk.END, f"保存 {table_name} 成功\n")
        scroll_to_end()  # 滚动到文本框的底部

        # 更新进度条
        current_progress = (i + 1) * progress_step
        update_progress(current_progress)

    # 关闭 ExcelWriter 对象以保存 Excel 文件
    excel_writer.close()
    header_excel_writer.close()
    new_excel_writer.close()

    text_output.insert(tk.END, "所有表格保存成功\n")

def show_save_selected_window():
    global text_output, save_path, original_data
    if not save_path:
        text_output.insert(tk.END, "请先保存主文件\n")
        return

    # 读取原始 Excel 文件
    original_file_path = save_path
    original_data = pd.read_excel(original_file_path, sheet_name=None)

    # 创建新的窗口
    top = tk.Toplevel()
    top.title("保存所选表格")

    # 输入要保存的表格编号部分
    tk.Label(top, text="输入要保存的表格编号（如1,2,5-10）：").pack(pady=5)
    table_numbers_entry = tk.Entry(top, width=30)
    table_numbers_entry.pack(pady=5)

    # 左右并列的文本框
    columns_frame = tk.Frame(top)
    columns_frame.pack(pady=5)

    all_columns_listbox = tk.Listbox(columns_frame, selectmode=tk.MULTIPLE, exportselection=0)
    all_columns_listbox.pack(side=tk.LEFT, padx=5)

    selected_columns_listbox = tk.Listbox(columns_frame, selectmode=tk.MULTIPLE, exportselection=0)
    selected_columns_listbox.pack(side=tk.LEFT, padx=5)

    # 添加原始表内的所有列标题到左侧的文本框
    all_columns = list(original_data[list(original_data.keys())[0]].columns)
    for column in all_columns:
        all_columns_listbox.insert(tk.END, column)

    # 按钮框架
    buttons_frame = tk.Frame(columns_frame)
    buttons_frame.pack(side=tk.LEFT, padx=5)

    def move_selected_to_right():
        selected_indices = all_columns_listbox.curselection()
        for index in selected_indices[::-1]:  # 从后向前移动，防止移除元素时索引变化
            selected_column = all_columns_listbox.get(index)
            selected_columns_listbox.insert(tk.END, selected_column)
            all_columns_listbox.delete(index)

    def move_selected_to_left():
        selected_indices = selected_columns_listbox.curselection()
        for index in selected_indices[::-1]:  # 从后向前移动，防止移除元素时索引变化
            selected_column = selected_columns_listbox.get(index)
            all_columns_listbox.insert(tk.END, selected_column)
            selected_columns_listbox.delete(index)

    def move_up():
        selected_indices = selected_columns_listbox.curselection()
        for index in selected_indices:
            if index == 0:  # 最上面的元素无法上移
                continue
            selected_column = selected_columns_listbox.get(index)
            selected_columns_listbox.delete(index)
            selected_columns_listbox.insert(index - 1, selected_column)
            selected_columns_listbox.select_set(index - 1)  # 选中移动后的元素

    def move_down():
        selected_indices = selected_columns_listbox.curselection()
        for index in selected_indices[::-1]:
            if index == selected_columns_listbox.size() - 1:  # 最下面的元素无法下移
                continue
            selected_column = selected_columns_listbox.get(index)
            selected_columns_listbox.delete(index)
            selected_columns_listbox.insert(index + 1, selected_column)
            selected_columns_listbox.select_set(index + 1)  # 选中移动后的元素

    move_to_right_button = tk.Button(buttons_frame, text="→", command=move_selected_to_right)
    move_to_right_button.pack(pady=5)

    move_to_left_button = tk.Button(buttons_frame, text="←", command=move_selected_to_left)
    move_to_left_button.pack(pady=5)

    move_up_button = tk.Button(buttons_frame, text="↑", command=move_up)
    move_up_button.pack(pady=5)

    move_down_button = tk.Button(buttons_frame, text="↓", command=move_down)
    move_down_button.pack(pady=5)

    def save_selected_columns():
        table_numbers_str = table_numbers_entry.get()
        selected_table_numbers = parse_table_numbers(table_numbers_str)

        if not selected_table_numbers:
            text_output.insert(tk.END, "请输入有效的表格编号\n")
            return

        selected_columns = [selected_columns_listbox.get(i) for i in range(selected_columns_listbox.size())]

        if not selected_columns:
            text_output.insert(tk.END, "请选择要保存的列\n")
            return

        # 创建一个 Pandas ExcelWriter 对象用于保存所选的表格数据
        selected_tables_save_path = save_path[:-5] + "_v3.xlsx"
        selected_tables_excel_writer = pd.ExcelWriter(selected_tables_save_path, engine='xlsxwriter')

        for table_number in selected_table_numbers:
            sheet_name = f"Table_{table_number}"
            if sheet_name in original_data:
                original_data[sheet_name][selected_columns].to_excel(selected_tables_excel_writer, sheet_name=sheet_name, index=False)
            else:
                text_output.insert(tk.END, f"表格 {sheet_name} 不存在\n")

        selected_tables_excel_writer.close()
        text_output.insert(tk.END, "所选表格保存成功\n")
        top.destroy()

    save_button = tk.Button(top, text="保存", command=save_selected_columns)
    save_button.pack(pady=5)

def parse_table_numbers(table_numbers_str):
    parts = table_numbers_str.replace(' ', '').split(',')
    table_numbers = []
    for part in parts:
        if '-' in part:
            start, end = part.split('-')
            try:
                start = int(start)
                end = int(end)
            except ValueError:
                return []  # 解析失败，返回空列表
            table_numbers.extend(range(start, end + 1))
        else:
            try:
                table_numbers.append(int(part))
            except ValueError:
                return []  # 解析失败，返回空列表
    return table_numbers

# 创建主窗口
root = tk.Tk()
root.title("打开文件并保存为Excel")

selected_file_label = tk.Label(root, text="选择的文件：")
selected_file_label.pack()

saved_file_label = tk.Label(root, text="保存的文件：")
saved_file_label.pack()

text_output = scrolledtext.ScrolledText(root, height=10, width=50)
text_output.pack()

progress_value = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_value, orient="horizontal", length=200, mode="determinate")
progress_bar.pack(pady=10)
progress_label = tk.Label(root, text="0.0%")
progress_label.pack()

open_button = tk.Button(root, text="打开文件,注意删除dat仪器校准部分", command=open_file)
open_button.pack(pady=10)

save_button = tk.Button(root, text="保存 Excel", command=save_excel)
save_button.pack(pady=10)

save_selected_button = tk.Button(root, text="保存所选表格", command=show_save_selected_window)
save_selected_button.pack(pady=10)

# 添加水印
watermark = tk.Label(root, text="© PyQ", fg="gray", font=("Arial", 10))
watermark.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)  # 在窗口右下角

root.mainloop()
