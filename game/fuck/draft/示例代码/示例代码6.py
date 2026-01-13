import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# 导入 openpyxl 的工具函数，用于将列索引转换为Excel列字母（A, B, C...）
from openpyxl.utils import get_column_letter

# 导入文件路径
excel_file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'
file_output_path = r'D:\Project\Numpy_Pandas\game\data\English logger.xlsx'

# 检查文件读取是否正确
try: 
    sheet_character = pd.read_excel(excel_file_path, sheet_name='角色活动祈愿')
except ValueError:
    print("警告: 无法找到'角色活动祈愿'相关数据或Excel文档损坏/被篡改。")
    print('请检查Excel文档是否损坏或被篡改。')
    exit()
except FileNotFoundError:
    print(f"错误: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
    exit()

# 检索需要转换的列名
column_to_convert = 'name'

# 列出需要转换的内容
list_map = {
    '琴': 'Jean',
}

sheet_map = {
    '角色活动祈愿': 'Character Event Prayers',
}

# 定义列宽
defined_column_widths = {
    'time': 21,
    'name': 25,
    'rarity': 10,
    'item_type': 15,
    'pity': 25,
    'total': 8, 
}

try:
    all_sheet_data = pd.read_excel(excel_file_path, sheet_name = None)
    converted_sheets = {}

    for original_sheet_name, df in all_sheet_data.items():
        new_sheet_name = sheet_map.get(original_sheet_name, original_sheet_name)

    
except FileNotFoundError:
    print(f"Error, the file '{excel_file_path}' not found.")
    print('Please check the file path and name.')
except KeyError as e:
    print(f"Data processing error,")
    print("The key is missing in the mapping dictionary or the column name does not exist.")
    print(f"More information: {e}")
except Exception as e:
    print(f"An unknown error occurred")
    print(f"More information: {e}")