import pandas as pd

def get_data_from_excel_pandas(file_path, sheet_name, col3_name, target_value_col3, col1_name, col2_name):
    """
    使用pandas在Excel的指定第三列查找固定值，并返回对应的第一列和第二列数据。

    Args:
        file_path (str): Excel文件的路径。
        sheet_name (str): 要处理的工作表名称。
        col3_name (str): 第三列的列名（例如 'Col C'）。
        target_value_col3: 要在第三列查找的固定数值。
        col1_name (str): 第一列的列名（例如 'Col A'）。
        col2_name (str): 第二列的列名（例如 'Col B'）。

    Returns:
        list: 包含字典的列表，每个字典包含匹配行的第一列和第二列数据。
    """
    try:
        # 读取Excel文件到DataFrame
        # header=0 表示第一行是列名。如果你的Excel没有列名，请设置为 header=None
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

        # 检查指定的列是否存在
        if col3_name not in df.columns:
            print(f"错误：列 '{col3_name}' 在Excel文件中不存在。")
            return []
        if col1_name not in df.columns:
            print(f"错误：列 '{col1_name}' 在Excel文件中不存在。")
            return []
        if col2_name not in df.columns:
            print(f"错误：列 '{col2_name}' 在Excel文件中不存在。")
            return []

        # 筛选出第三列值等于目标值的行
        # 使用 .astype(str) 确保类型一致性，避免比较问题
        # .fillna('') 将 NaN 转换为空字符串，防止比较 NoneType
        filtered_df = df[df[col3_name].astype(str).fillna('') == str(target_value_col3)]

        # 如果没有找到匹配项
        if filtered_df.empty:
            print(f"在 '{col3_name}' 列中未找到 '{target_value_col3}'。")
            return []

        # 提取第一列和第二列的数据，并转换为字典列表
        # 'records' 格式将每一行转换为一个字典
        results = filtered_df[[col1_name, col2_name]].to_dict(orient='records')
        
        # 如果需要包含原始行号，可以使用 df.index
        # 此时需要将 df 的索引重置为原始行号（如果Excel有头部，则需要+2）
        # 或者直接使用 filtered_df.index 获取匹配行的索引（0-based）
        # 示例：
        # results_with_index = []
        # for index, row_data in filtered_df.iterrows():
        #     results_with_index.append({
        #         "行号": index + 2, # +2 是因为 header=0 会使 DataFrame 索引从0开始，且跳过了Excel的第一行标题
        #         col1_name: row_data[col1_name],
        #         col2_name: row_data[col2_name]
        #     })
        # return results_with_index
        
        return results

    except FileNotFoundError:
        print(f"错误：文件 '{file_path}' 未找到。")
    except ValueError as e:
        print(f"错误：读取Excel文件时发生问题，可能是工作表名称错误或文件格式问题: {e}")
    except Exception as e:
        print(f"发生未知错误: {e}")
    
    return []


# --- 使用示例 ---
excel_file = 'your_excel_file.xlsx'  # 替换为你的Excel文件路径
sheet_name = 'Sheet1'                 # 替换为你的工作表名称
target_number = 100                   # 替换为你要查找的固定数值

# 注意：这里需要提供列的实际名称，而不是数字索引
col_c_name = 'Col C'  # 你的第三列的标题
col_a_name = 'Col A'  # 你的第一列的标题
col_b_name = 'Col B'  # 你的第二列的标题

found_data_pandas = get_data_from_excel_pandas(
    excel_file, sheet_name, col_c_name, target_number, col_a_name, col_b_name
)

if found_data_pandas:
    print(f"\n使用 pandas 找到 '{target_number}' 的数据：")
    for item in found_data_pandas:
        print(f"  第一列: {item[col_a_name]}, 第二列: {item[col_b_name]}")
else:
    print(f"使用 pandas 在第三列中未找到 '{target_number}'。")



# 代码解释：
# pd.read_excel(file_path, sheet_name=sheet_name, header=0): 读取 Excel 文件到 DataFrame。
# header=0 告诉 pandas 你的 Excel 文件的第一行是列名（索引为0）。如果你的文件没有列名，请使用 header=None，此时列名会默认是 0, 1, 2...。
# df[col3_name].astype(str).fillna('') == str(target_value_col3):
# df[col3_name] 选择名为 col3_name 的那一列。
# .astype(str) 将该列的所有值转换为字符串。
# .fillna('') 填充 NaN (Not a Number，缺失值) 为空字符串，防止比较时出现错误。
# == str(target_value_col3) 进行比较。这会返回一个布尔序列（True/False），用于筛选行。
# filtered_df = df[...]: 使用布尔序列来筛选 DataFrame，只保留满足条件的行。
# filtered_df[[col1_name, col2_name]]: 从筛选后的 DataFrame 中选择你需要的列。
# .to_dict(orient='records'): 将筛选并选择后的 DataFrame 转换为一个字典列表，每个字典代表一行数据。'records' 参数表示每个字典的键是列名，值是对应单元格的数据。