import pandas as pd
# import matplotlib.pyplot as plt
# import seaborn as sns
# import numpy as np # 如果没有用到 np.nan，此行可以移除
# import os

# 导入 openpyxl 的工具函数，用于将列索引转换为Excel列字母（A, B, C...）
from openpyxl.utils import get_column_letter

# 配置excel相关文件
excel_file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'
file_output_path = r'D:\Project\Numpy_Pandas\game\data\English logger.xlsx'

# 初始的df_character读取如果仅用于检查文件存在性或某个特定表的格式，可以保留。
try:
    test_df_character = pd.read_excel(excel_file_path, sheet_name='角色活动祈愿')
except ValueError:
    print("警告: 无法找到'角色活动祈愿'相关数据或Excel文档损坏/被篡改。")
    print('请检查Excel文档是否损坏或被篡改。')
    exit()
except FileNotFoundError:
    print(f"错误: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
    exit()


column_to_convert = 'name'

list_map = {
    '琴': 'Jean',
    # '迪卢克': 'Diluc',
    # '温迪': 'Venti',
    # '可莉': 'Klee',
    # '莫娜': 'Mona',
    # '阿贝多': 'Albedo',
    # '优菈': 'Eula',
    # '魈': 'Xiao',
    # '刻晴': 'Keqing',
    # '七七': 'Qiqi',
    # '达达利亚': 'Tartaglia',
    # '钟离': 'Zhongli',
    # '甘雨': 'Ganyu',
    # '胡桃': 'Hu Tao',
    # '申鹤': 'Shenhe',
    # '夜兰': 'Yelan',
    # '白术': 'Baizhu',
    # '闲云': 'Xianyun',
    # '神里绫华': 'Kamisato Ayaka',
    # '枫原万叶': 'Kaedehara Kazuha',
    # '宵宫': 'Yoimiya',
    # '雷电将军': 'Raiden Shogun',
    # '珊瑚宫心海': 'Sangonomiya Kokomi',
    # '荒泷一斗': 'Arataki Itto',
    # '八重神子': 'Yae Miko',
    # '神里绫人': 'Kamisato Ayato',
    # '梦见月瑞希': 'Yumemizuki Mizuki',
    # '提纳里': 'Tighnari',
    # '赛诺': 'Cyno',
    # '妮露': 'Nilou',
    # '纳西妲': 'Nahida',
    # '流浪者': 'Wanderer',
    # '艾尔海森': 'Alhaitham',
    # '迪希雅': 'Dehya',
    # '林尼': 'Lyney',
    # '那维莱特': 'Neuvillette',
    # '莱欧斯利': 'Wriothesley',
    # '芙宁娜': 'Furina',
    # '纳维娅': 'Navia',
    # '千织': 'Chiori',
    # '阿蕾奇诺': 'Arlecchino',
    # '克洛琳德': 'Clorinde',
    # '希格雯': 'Sigewinne',
    # '艾梅莉埃': 'Emilie',
    # '爱可菲': 'Escoffier',
    # '玛拉妮': 'Mualani',
    # '基尼奇': 'Kinich',
    # '希诺宁': 'Xilonen',
    # '恰斯卡': 'Chasca',
    # '茜特拉莉': 'Citlali',
    # '玛薇卡': 'Mavuika',
    # '瓦雷莎': 'Varesa',
    # '伊法': 'Ifa',
    # '丝柯克': 'Skirk',
    # '伊涅夫': 'Ineffa',
}
sheet_map = {
    '角色活动祈愿': 'Character Event Prayers',
    # '武器活动祈愿': 'Weapon Event Prayer',
    # '常驻祈愿': 'Permanent Prayer',
    # '集录祈愿': 'Collection of Prayers',
    # '新手祈愿': "Novice's Prayer",
}

# --- 新增：定义列宽 ---
# 这是一个字典，键是列名，值是该列期望的宽度。
# 请根据你的实际Excel文件中存在的列名以及你希望的宽度来调整这里。
# 如果某个列名没有在这里定义，它将保持Excel的默认宽度。
defined_column_widths = {
    'time': 21,       # 例如：时间列宽度设为20
    'name': 25,       # 例如：名字列宽度设为25
    'rarity': 10,     # 例如：稀有度列宽度设为10
    'item_type': 15,  # 例如：物品类型列宽度设为15
    'pity': 25,        # 垫刀数
    'total': 8,       # 总抽数
    # 可以继续添加其他列
}
# --- 结束新增 ---

try:
    all_sheet_data = pd.read_excel(excel_file_path, sheet_name=None)
    converted_sheets = {}

    for original_sheet_name, df in all_sheet_data.items():
        new_sheet_name = sheet_map.get(original_sheet_name, original_sheet_name)

        if column_to_convert in df.columns:
            df_filtered = df[df[column_to_convert].isin(list_map.keys())].copy()
            df_filtered.loc[:, column_to_convert] = df_filtered[column_to_convert].replace(list_map)
            converted_sheets[new_sheet_name] = df_filtered
        else:
            print(f"警告: 工作表 '{original_sheet_name}' 中不存在列 '{column_to_convert}'，跳过该列的数据转换和行删除。将保留原始数据。")
            converted_sheets[new_sheet_name] = df # 保持原始数据

    with pd.ExcelWriter(file_output_path, engine='openpyxl') as writer:
        for sheet_name, df_to_writer in converted_sheets.items():
            df_to_writer.to_excel(writer, sheet_name=sheet_name, index=False)

            # --- 新增：设置列宽 ---
            # 获取当前写入的工作表对象
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # 遍历 DataFrame 的列，并根据定义的宽度设置Excel列宽
            for i, col_name in enumerate(df_to_writer.columns):
                # DataFrame的列索引是0开始的，但Excel的列是从1开始的（A=1, B=2...）
                # 因此需要 i + 1
                col_letter = get_column_letter(i + 1)
                
                # 检查这个列名是否在我们的宽度定义字典中
                if col_name in defined_column_widths:
                    desired_width = defined_column_widths[col_name]
                    # 设置列的宽度
                    worksheet.column_dimensions[col_letter].width = desired_width
                # else:
                    # 如果没有定义，列将保持默认宽度，可以根据需要添加日志或默认设置
            # --- 结束新增 ---

    print(f"数据已成功转换并保存到 '{file_output_path}'")

except FileNotFoundError:
    print(f"错误: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
except KeyError as e:
    print(f"数据处理错误: 映射字典中缺少键或列名不存在。详细信息: {e}")
except Exception as e:
    print(f"发生未知错误: {e}")