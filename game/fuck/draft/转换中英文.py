import pandas as pd

# --- 配置部分 ---
# 输入Excel文件路径
input_excel_file = 'input.xlsx'
# 输出Excel文件路径 (可以与输入文件不同，避免覆盖)
output_excel_file = 'output_converted_with_sheets.xlsx'

# 定义需要进行转换的列名
column_to_convert = '状态' # 根据你的Excel列名修改

# 定义中文到英文的映射字典 (针对列数据)
conversion_map_column_data = {
    '待处理': 'Pending',
    '已完成': 'Completed',
    '进行中': 'In Progress',
    '未开始': 'Not Started',
    '审核中': 'Under Review',
    '已拒绝': 'Rejected'
    # 添加更多你需要的映射
}

# 定义中文到英文的映射字典 (针对Sheet名称)
conversion_map_sheet_names = {
    '销售数据': 'Sales Data',
    '库存信息': 'Inventory Info',
    '用户列表': 'User List',
    '报表': 'Report',
    # 添加更多你需要的Sheet名称映射
}

# --- 代码逻辑部分 ---
try:
    # 1. 读取Excel文件中的所有Sheet到DataFrame字典
    # sheet_name=None 会返回一个字典，键是Sheet名称，值是对应的DataFrame
    all_sheets_data = pd.read_excel(input_excel_file, sheet_name=None)
    print(f"成功读取文件: '{input_excel_file}'")
    print(f"原始Sheet名称: {list(all_sheets_data.keys())}")

    converted_sheets_for_writing = {}

    # 2. 遍历每个Sheet，进行数据和名称转换
    for original_sheet_name, df in all_sheets_data.items():
        print(f"\n--- 处理 Sheet: '{original_sheet_name}' ---")

        # 2.1 获取新的Sheet名称
        # 如果原始Sheet名称在映射字典中，则使用映射后的名称，否则保持不变
        new_sheet_name = conversion_map_sheet_names.get(original_sheet_name, original_sheet_name)
        if new_sheet_name != original_sheet_name:
            print(f"Sheet名称 '{original_sheet_name}' 转换为 '{new_sheet_name}'")
        else:
            print(f"Sheet名称 '{original_sheet_name}' 无需转换。")

        # 2.2 对指定列进行中文到英文的转换 (如果该列存在于当前Sheet中)
        if column_to_convert in df.columns:
            # 打印原始数据的前几行（如果需要）
            # print(f"Sheet '{original_sheet_name}' 列 '{column_to_convert}' 原始数据:")
            # print(df[column_to_convert].head())

            df[column_to_convert] = df[column_to_convert].replace(conversion_map_column_data)
            print(f"列 '{column_to_convert}' 在此Sheet中已完成转换。")

            # 打印转换后的数据的前几行（如果需要）
            # print(f"Sheet '{original_sheet_name}' 列 '{column_to_convert}' 转换后数据:")
            # print(df[column_to_convert].head())
        else:
            print(f"警告: 列 '{column_to_convert}' 不存在于 Sheet '{original_sheet_name}' 中，跳过列数据转换。")

        # 将处理后的DataFrame和新的Sheet名称存储起来，以便后续写入
        converted_sheets_for_writing[new_sheet_name] = df

    # 3. 使用ExcelWriter将所有处理过的DataFrame写入新的Excel文件
    # engine='openpyxl' 是用于处理 .xlsx 文件的后端引擎
    with pd.ExcelWriter(output_excel_file, engine='openpyxl') as writer:
        for sheet_name, df_to_write in converted_sheets_for_writing.items():
            df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Sheet '{sheet_name}' 已写入。")

    print(f"\n所有转换已完成。结果已保存到: '{output_excel_file}'")

except FileNotFoundError:
    print(f"错误: 找不到文件 '{input_excel_file}'。请检查文件路径和名称是否正确。")
except KeyError as e:
    print(f"数据处理错误: {e}")
except Exception as e:
    print(f"发生未知错误: {e}")