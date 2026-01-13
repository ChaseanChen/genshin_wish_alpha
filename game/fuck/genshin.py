import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os

# --- 配置区 ---
# 请替换为你的Excel文件路径
EXCEL_FILE_PATH = r'D:\Project\Numpy_Pandas\game\data\Genshin Wishs.xlsx'

# 根据你提供的Excel截图，设置正确的列名顺序。
# '保底内' 看起来是冗余的，但保留它以匹配你的文件结构。
# '祈愿类型' 在你的截图中缺失，脚本将通过'总次数'的重置来自动识别Pity序列。
COLUMN_NAMES = ['时间', '名称', '类别', '星级', '总次数', '保底内', '备注']

# 可以统一的物品类别名称映射（如果你的Excel中有多种叫法，方便归类统计）
# 例如: '角色': '角色', '武器': '武器'
ITEM_TYPE_MAPPING = {
    '角色': '角色',
    '武器': '武器'
}
# --- 配置区结束 ---


# 设置图表显示中文 (MacOS和Linux用户可能需要安装对应字体)
plt.rcParams['font.sans-serif'] = ['SimHei', 'FangSong', 'STSong', 'Microsoft YaHei', 'Arial Unicode MS'] # 尝试多种字体
plt.rcParams['axes.unicode_minus'] = False # 解决保存图像时负号'-'显示为方块的问题

print("--- 原神祈愿数据分析脚本 (针对您提供的Excel格式) ---")

# --- 1. 数据载入与初步清洗 ---
print("\n--- 1. 数据载入与初步清洗 ---")
df = pd.DataFrame() # 初始化一个空DataFrame，防止文件未找到时报错

if not os.path.exists(EXCEL_FILE_PATH):
    print(f"错误：文件未找到。请检查文件路径和名称是否正确: {EXCEL_FILE_PATH}")
    exit()

try:
    # 指定 header=None 表示没有标题行，或者如果第一行是标题，则指定 header=0
    # 但是根据图片，第一行是标题，所以默认header=0即可
    df = pd.read_excel(EXCEL_FILE_PATH)
    print("数据载入成功！")
    print(f"原始数据维度: {df.shape}")
    # print(df.head()) # 可以取消注释查看原始数据前几行
except Exception as e:
    print(f"载入Excel文件时发生错误: {e}")
    print("请确认Excel文件格式正确，且未被其他程序占用。")
    exit()

# 统一列名
if len(df.columns) == len(COLUMN_NAMES):
    df.columns = COLUMN_NAMES
    print(f"列名已设置为: {COLUMN_NAMES}")
else:
    print(f"警告: Excel文件列数 ({len(df.columns)}) 与预设列数 ({len(COLUMN_NAMES)}) 不匹配。")
    print("请检查 COLUMN_NAMES 配置是否正确，或者手动调整df.columns。")
    print(f"当前Excel列名: {df.columns.tolist()}")
    print("如果列名不匹配，后续操作可能出错。脚本将尝试继续...")


# 将时间列转换为datetime类型，并按时间排序 (非常重要，Pity计算依赖顺序)
try:
    df['时间'] = pd.to_datetime(df['时间'])
    df = df.sort_values(by='时间').reset_index(drop=True)
    print("时间列转换成功并已排序。")
except KeyError:
    print("错误：'时间'列未找到，请检查 COLUMN_NAMES 配置。")
    exit()
except Exception as e:
    print(f"将'时间'列转换为日期时间时发生错误: {e}")
    exit()

# 检查缺失值或异常值
print("\n缺失值检查:")
print(df.isnull().sum())
# 可以根据需要处理缺失值，例如：df.dropna(subset=['星级', '总次数'], inplace=True)

# 确保星级和总次数是数值类型
try:
    df['星级'] = df['星级'].astype(int)
    print("星级列已转换为整数类型。")
except KeyError:
    print("错误：'星级'列未找到，请检查 COLUMN_NAMES 配置。")
    exit()
except Exception as e:
    print(f"将'星级'列转换为整数时发生错误: {e}")
    print("请确认'星级'列中只包含数字。")
    exit()

try:
    df['总次数'] = df['总次数'].astype(int)
    print("总次数列已转换为整数类型。")
except KeyError:
    print("错误：'总次数'列未找到，请检查 COLUMN_NAMES 配置。")
    exit()
except Exception as e:
    print(f"将'总次数'列转换为整数时发生错误: {e}")
    print("请确认'总次数'列中只包含数字。")
    exit()

# 统一物品类别名称
if ITEM_TYPE_MAPPING:
    df['类别'] = df['类别'].replace(ITEM_TYPE_MAPPING)
    print("物品类别已统一。")

print("\n清洗后数据前5行:")
print(df.head())
print("\n数据信息:")
print(df.info())

# --- 2. 核心处理：识别 Pity 序列并补充信息 ---
print("\n--- 2. 核心处理：识别 Pity 序列并补充信息 ---")

# 识别 Pity 序列：当 '总次数' 为 1 时，表示一个新的 Pity 序列开始
df['is_new_sequence'] = (df['总次数'] == 1)
df['pity_sequence_id'] = df['is_new_sequence'].cumsum()

# Pity 值即为 '总次数'
df.rename(columns={'总次数': 'Pity'}, inplace=True)

# 处理每个 Pity 序列，追踪上次5星物品和判断是否歪
df_final = pd.DataFrame()

# 仅用于跟踪上一个序列的最后5星物品，以便于判断跨序列的“歪”？
# 实际上，原神的保底机制是针对**同类祈愿池**生效的。
# 由于我们没有“祈愿类型”列，所以无法准确判断大小保底和是否歪。
# 这里的'上次5星物品'和'是否歪'将仅在当前识别的'pity_sequence_id'内进行跟踪。
# 如果你想更精确，需要补充'祈愿类型'信息或手动判断。
def process_pity_sequence(df_sequence):
    df_sequence = df_sequence.copy()
    df_sequence['上次5星物品'] = ''
    df_sequence['是否歪'] = '待手动判断' # 无法自动判断，需要知道UP角色/武器
    # df_sequence['是否为UP'] = '待手动判断' # 同上

    last_5_star_item_in_sequence = ''

    for i, row in df_sequence.iterrows():
        df_sequence.at[i, '上次5星物品'] = last_5_star_item_in_sequence

        if row['星级'] == 5:
            # 5星出货，更新上次5星物品
            last_5_star_item_in_sequence = row['名称']
            # 这里是判断是否歪的关键，需要外部UP信息
            # 例如： if row['名称'] not in get_current_up_items(row['时间'], row['类别']):
            #          df_sequence.at[i, '是否歪'] = '是'
            #        else:
            #          df_sequence.at[i, '是否歪'] = '否'
            pass # 暂时无法自动判断是否歪

    return df_sequence

unique_sequences = df['pity_sequence_id'].unique()
if len(unique_sequences) == 0:
    print("警告：未能识别到任何Pity序列，请检查数据。")
else:
    for seq_id in unique_sequences:
        print(f"正在处理 Pity 序列 ID: {seq_id}...")
        df_sequence = df[df['pity_sequence_id'] == seq_id].copy()
        df_final = pd.concat([df_final, process_pity_sequence(df_sequence)])

    df_final = df_final.sort_values(by='时间').reset_index(drop=True)
    print("\nPity序列处理完成！结果前5行:")
    print(df_final.head())
    print("\n总祈愿次数（Pity计算后）:", len(df_final))

print("\n--- 分析局限性说明 ---")
print("由于Excel文件中缺少“祈愿类型”列（如：角色活动祈愿、武器活动祈愿、常驻祈愿），")
print("本脚本无法区分不同的祈愿池类型。所有分析均基于'总次数'列自动识别的Pity序列。")
print("这意味着：")
print("1. 无法按官方祈愿池类型（角色池、武器池等）进行单独统计和分析。")
print("2. '是否歪'的判断是无法自动进行的，因为脚本不知道当时的UP角色/武器是谁。")
print("如果需要更深入的分析，建议获取包含'祈愿类型'的完整祈愿数据。")
print("-----------------------")


# --- 3. 基础统计分析 ---
print("\n--- 3. 基础统计分析 ---")

if df_final.empty:
    print("无数据可供分析，请检查数据载入和Pity序列处理步骤。")
else:
    total_pulls = len(df_final)
    print(f"\n总祈愿次数: {total_pulls}")

    # 1. 各星级出货数量与概率
    rarity_counts = df_final['星级'].value_counts().sort_index(ascending=False)
    rarity_percentages = df_final['星级'].value_counts(normalize=True).sort_index(ascending=False) * 100

    print("\n各星级出货数量:")
    print(rarity_counts)
    print("\n各星级出货概率:")
    print(rarity_percentages.round(2))
    print(f"(官方基础5星概率: 0.6%, 4星概率: 5.1%)")

    # 2. 各Pity序列出货数量与概率
    print("\n各Pity序列出货统计:")
    # 筛选出实际出货的5星和4星的Pity值
    pity_5_star_pulls = df_final[df_final['星级'] == 5]
    pity_4_star_pulls = df_final[df_final['星级'] == 4]

    # 计算平均Pity (这里Pity就是总次数，已经是从1开始计数)
    avg_5_star_pity = pity_5_star_pulls['Pity'].mean() if not pity_5_star_pulls.empty else 0
    avg_4_star_pity = pity_4_star_pulls['Pity'].mean() if not pity_4_star_pulls.empty else 0
    print(f"\n平均5星Pity (即'总次数'在5星出货时的平均值): {avg_5_star_pity:.2f}")
    print(f"平均4星Pity (即'总次数'在4星出货时的平均值): {avg_4_star_pity:.2f}")

    # 各Pity序列下的平均Pity
    print("\n各Pity序列下的平均Pity:")
    pity_by_sequence_5 = pity_5_star_pulls.groupby('pity_sequence_id')['Pity'].mean()
    print("5星平均Pity:")
    print(pity_by_sequence_5.round(2))

    pity_by_sequence_4 = pity_4_star_pulls.groupby('pity_sequence_id')['Pity'].mean()
    print("\n4星平均Pity:")
    print(pity_by_sequence_4.round(2))

    # 3. 最长未出货记录 (即最大Pity值)
    max_5_star_pity = pity_5_star_pulls['Pity'].max() if not pity_5_star_pulls.empty else 0
    max_4_star_pity = pity_4_star_pulls['Pity'].max() if not pity_4_star_pulls.empty else 0
    print(f"\n最长5星Pity (单次序列内): {max_5_star_pity} 抽")
    print(f"最长4星Pity (单次序列内): {max_4_star_pity} 抽")

    # 4. 具体物品统计（例如，抽到哪些5星）
    print("\n抽到的所有5星物品列表:")
    five_star_items = df_final[df_final['星级'] == 5][['名称', '类别', 'Pity', '时间', '上次5星物品', 'pity_sequence_id']].sort_values(by='时间')
    if not five_star_items.empty:
        print(five_star_items.to_string(index=False))
    else:
        print("暂未抽到5星物品。")


    # --- 4. 高级分析与可视化 ---
    print("\n--- 4. 高级分析与可视化 ---")

    # 设置Seaborn风格
    sns.set_style("whitegrid")

    # 4.1 5星Pity分布直方图 (软保底分析)
    plt.figure(figsize=(12, 7))
    if not pity_5_star_pulls.empty:
        sns.histplot(pity_5_star_pulls['Pity'], bins=range(1, 91), kde=True)
        plt.title('5星Pity分布')
        plt.xlabel('Pity (抽数)')
        plt.ylabel('出货次数')
        plt.axvline(x=74, color='r', linestyle='--', label='官方软保底开始 (74抽)')
        plt.legend()
    else:
        plt.text(0.5, 0.5, '暂无5星数据进行Pity分布分析', horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
    plt.grid(axis='y', linestyle='--')
    plt.tight_layout()
    plt.show()

    # 4.2 4星Pity分布直方图 (软保底分析)
    plt.figure(figsize=(12, 7))
    if not pity_4_star_pulls.empty:
        sns.histplot(pity_4_star_pulls['Pity'], bins=range(1, 11), kde=True)
        plt.title('4星Pity分布')
        plt.xlabel('Pity (抽数)')
        plt.ylabel('出货次数')
        plt.axvline(x=9, color='r', linestyle='--', label='官方软保底开始 (9抽)')
        plt.legend()
    else:
        plt.text(0.5, 0.5, '暂无4星数据进行Pity分布分析', horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
    plt.grid(axis='y', linestyle='--')
    plt.tight_layout()
    plt.show()

    # 4.3 总星级分布饼图
    plt.figure(figsize=(8, 8))
    if not rarity_counts.empty:
        plt.pie(rarity_counts, labels=[f'{idx}星 ({count})' for idx, count in rarity_counts.items()],
                autopct='%1.1f%%', startangle=90, colors=sns.color_palette("pastel"))
        plt.title('总祈愿星级分布')
        plt.axis('equal') # 确保饼图是圆形的
    else:
        plt.text(0.5, 0.5, '无数据可用于星级分布饼图', horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
    plt.tight_layout()
    plt.show()

    # 4.4 各Pity序列星级出货概率条形图
    plt.figure(figsize=(14, 8))
    # 按照pity_sequence_id和星级分组计算百分比
    sequence_summary_percentage = df_final.groupby('pity_sequence_id')['星级'].value_counts(normalize=True).unstack(fill_value=0) * 100

    if not sequence_summary_percentage.empty:
        # 重塑数据以便于绘制堆叠条形图
        sequence_summary_percentage_plot = sequence_summary_percentage.stack().reset_index()
        sequence_summary_percentage_plot.columns = ['Pity序列ID', '星级', '概率']
        sns.barplot(data=sequence_summary_percentage_plot, x='Pity序列ID', y='概率', hue='星级',
                    palette='viridis', dodge=True)
        plt.title('各Pity序列星级出货概率')
        plt.ylabel('概率 (%)')
        plt.xlabel('Pity序列ID')
        plt.xticks(rotation=45, ha='right')
        plt.legend(title='星级')
        plt.grid(axis='y', linestyle='--')
    else:
        plt.text(0.5, 0.5, '无数据可用于Pity序列出货概率图', horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
    plt.tight_layout()
    plt.show()

    # 4.5 5星Pity累计分布函数 (CDF)
    plt.figure(figsize=(12, 7))
    if not pity_5_star_pulls.empty:
        sns.ecdfplot(data=pity_5_star_pulls, x='Pity', label='5星Pity')
        plt.title('5星Pity累计分布函数')
        plt.xlabel('Pity (抽数)')
        plt.ylabel('累计概率 (小于等于当前Pity的概率)')
        plt.grid(True)
        plt.legend()
    else:
        plt.text(0.5, 0.5, '暂无5星数据进行CDF分析', horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
    plt.tight_layout()
    plt.show()

    print("\n--- 分析完成 ---")
    print("请查看弹出的图表窗口和命令行输出。")