import pandas as pd
import openpyxl

# --- 配置相关文件地址 --- #
input_file_path = r'D:\Project\Pydata\game\data\原神祈愿记录.xlsx'
output_file_path = r'D:\Project\Pydata\game\data\New file.xlsx'

column_to_convert = '星级'

# --- 代码逻辑部分 --- #
try:
    all_sheets_data = pd.read_excel(input_file_path, sheet_name = None)
    print(f"成功读取文件内容: {input_file_path}")
    print(f"初始Sheet名称: {list(all_sheets_data.keys())}")

    converted_sheet_for_writing = {}

    for original_sheet_name, df in all_sheets_data.items():
        print(f"\n --- 处理 Sheet: {original_sheet_name} --- ")

except FileNotFoundError:
    print(f"错误: 找不到文件: {input_file_path}")
    print('请检查文件路径和名称是否正确')

except KeyError as e:
    print(f"数据处理错误: {e}")

except Exception as e:
    print(f"发生未知错误: {e}")
