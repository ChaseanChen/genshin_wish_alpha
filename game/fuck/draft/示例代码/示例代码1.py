import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 获取文档整体内容
file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'
file_output_path = r'D:\Project\Numpy_Pandas\game\data\English logger.xlsx'

# 检测文档角色活动祈愿页读取是否正确
try:
    character_original_data = pd.read_excel(file_path, sheet_name = '角色活动祈愿')

except ValueError:
    print("警告: 无法找到'角色活动祈愿'相关数据")
    print('请检查Excel文档是否损坏或被篡改')
    print("Warning: Unable to find data related to 'Character Activity Prayer'")
    print("Please check whether the Excel document is damaged or tampered with")
    exit()

# 获取角色活动祈愿页面的DataFrame
df_character_original_data = pd.DataFrame(character_original_data)

# 获取5星角色行
five_character_stars = df_character_original_data[df_character_original_data['rarity'] == 5].copy()

title_to_convert = 'name'

name_list_map = {
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

# try:
    # all_sheet_data = pd.read_excel(df_character_original_data, sheet_name = None)
    # converted_sheets = {}
    