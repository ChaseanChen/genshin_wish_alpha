import pandas as pd
# import matplotlib.pyplot as plt # 如果不需要绘图，可以移除
# import seaborn as sns           # 如果不需要绘图，可以移除
# import numpy as np              # 如果没有用到数值操作，可以移除
# import os

# 配置excel相关文件
excel_file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'
file_output_path = r'D:\Project\Numpy_Pandas\game\data\English logger.xlsx'

# 初始的df_character读取如果仅用于检查文件存在性或某个特定表的格式，可以保留。
# 但如果只是为了转换所有表，下面的主读取循环就足够了。
# 暂时保留以匹配原代码意图，但它在功能上并非必须。
try:
    # 这一步仅用于检查文件是否存在以及'角色活动祈愿'表是否可读
    # 如果文件不存在，下面的主读取会捕获 FileNotFoundError
    test_df_character = pd.read_excel(excel_file_path, sheet_name='角色活动祈愿')
    # df_character_original = pd.DataFrame(test_df_character) # 这行也是冗余的，如果df_character_original不再被使用
    # five_character_stars = df_character_original[df_character_original['rarity'] == 5].copy() # 这行是未使用的
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
    '迪卢克': 'Diluc',
    '温迪': 'Venti',
    '可莉': 'Klee', # 小写k统一为大写K
    '莫娜': 'Mona',
    '阿贝多': 'Albedo',
    '优菈': 'Eula',
    '魈': 'Xiao',
    '刻晴': 'Keqing',
    '七七': 'Qiqi',
    '达达利亚': 'Tartaglia',
    '钟离': 'Zhongli',
    '甘雨': 'Ganyu',
    '胡桃': 'Hu Tao',
    '申鹤': 'Shenhe',
    '夜兰': 'Yelan',
    '白术': 'Baizhu',
    '闲云': 'Xianyun',
    '神里绫华': 'Kamisato Ayaka',
    '枫原万叶': 'Kaedehara Kazuha',
    '宵宫': 'Yoimiya',
    '雷电将军': 'Raiden Shogun',
    '珊瑚宫心海': 'Sangonomiya Kokomi',
    '荒泷一斗': 'Arataki Itto',
    '八重神子': 'Yae Miko',
    '神里绫人': 'Kamisato Ayato',
    '梦见月瑞希': 'Yumemizuki Mizuki',
    '提纳里': 'Tighnari',
    '赛诺': 'Cyno',
    '妮露': 'Nilou',
    '纳西妲': 'Nahida',
    '流浪者': 'Wanderer',
    '艾尔海森': 'Alhaitham',
    '迪希雅': 'Dehya',
    '林尼': 'Lyney',
    '那维莱特': 'Neuvillette',
    '莱欧斯利': 'Wriothesley', # 修正拼写
    '芙宁娜': 'Furina',
    '纳维娅': 'Navia',
    '千织': 'Chiori',
    '阿蕾奇诺': 'Arlecchino',
    '克洛琳德': 'Clorinde',
    '希格雯': 'Sigewinne',
    '艾梅莉埃': 'Emilie',
    '爱可菲': 'Escoffier',
    '玛拉妮': 'Mualani',
    '基尼奇': 'Kinich',
    '希诺宁': 'Xilonen',
    '恰斯卡': 'Chasca',
    '茜特拉莉': 'Citlali',
    '玛薇卡': 'Mavuika',
    '瓦雷莎': 'Varesa',
    '伊法': 'Ifa',
    '丝柯克': 'Skirk',
    '伊涅夫': 'Ineffa',
}
sheet_map = {
    '角色活动祈愿': 'Character Event Prayers',
    '武器活动祈愿': 'Weapon Event Prayer',
    '常驻祈愿': 'Permanent Prayer',
    '集录祈愿': 'Collection of Prayers',
    '新手祈愿': "Novice's Prayer",
}

try:
    # 修正1: pd.read_excel() 的第一个参数应该是文件路径
    all_sheet_data = pd.read_excel(excel_file_path, sheet_name=None)
    converted_sheets = {}

    for original_sheet_name, df in all_sheet_data.items():
        new_sheet_name = sheet_map.get(original_sheet_name, original_sheet_name)

        if column_to_convert in df.columns:
            # 使用 .loc 进行赋值以避免 SettingWithCopyWarning
            df.loc[:, column_to_convert] = df[column_to_convert].replace(list_map)
        else:
            print(f"警告: 工作表 '{original_sheet_name}' 中不存在列 '{column_to_convert}'，跳过该列的数据转换。")

        # 修正2: 这行必须在 for 循环内部，以确保所有工作表都被添加
        converted_sheets[new_sheet_name] = df

    with pd.ExcelWriter(file_output_path, engine='openpyxl') as writer:
        for sheet_name, df_to_writer in converted_sheets.items():
            df_to_writer.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"数据已成功转换并保存到 '{file_output_path}'")

except FileNotFoundError:
    print(f"错误: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
except KeyError as e:
    print(f"数据处理错误: 映射字典中缺少键或列名不存在。详细信息: {e}")
except Exception as e:
    print(f"发生未知错误: {e}")