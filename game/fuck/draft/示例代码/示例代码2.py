import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
# import os

#配置excel相关文件
excel_file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'
file_output_path = r'D:\Project\Numpy_Pandas\game\data\English logger.xlsx'

try:
    df_character = pd.read_excel(excel_file_path, sheet_name = '角色活动祈愿')
except ValueError:
    print("警告: 无法找到'角色活动祈愿'相关数据")
    print('请检查Excel文档是否损坏或被篡改')
    exit()

df_character_original = pd.DataFrame(df_character)
five_character_stars = df_character_original[df_character_original['rarity'] == 5].copy()

column_to_convert = 'name'

list_map = {
    '琴': 'Jean',
    '迪卢克': 'Diluc',
    '温迪': 'Venti',
    '可莉': 'klee',
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
    '莱欧斯利': 'Wirithesley',
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
    all_sheet_data = pd.read_excel(df_character_original, sheet_name = None)
    converted_sheets = {}
    for original_sheet_name, df in all_sheet_data.items():
        new_sheet_name = sheet_map.get(original_sheet_name, original_sheet_name)
    
    if column_to_convert in df.columns:
        df[column_to_convert] = df[column_to_convert].replace(list_map)
    else:
        print(f"Warning: Column '{column_to_convert}' does not exist in Sheet '{original_sheet_name}', skipping column data conversion.")

        converted_sheets[new_sheet_name] = df

    with pd.ExcelWriter(file_output_path, engine = 'openpyxl') as writer:
        for sheet_name, df_to_writer in converted_sheets.items():
            df_to_writer.to_excel(writer, sheet_name = sheet_name, index = False)

except FileNotFoundError:
    print(f"Error: File '{excel_file_path}' not found. Please check that the file path and name are correct.")
except KeyError as e:
    print(f"Data processing error: {e}")
except Exception as e:
    print(f"An unknown error occurred: {e}")