import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os

# 配置相关文件内容
excel_file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'

try:
    df_character = pd.read_excel(excel_file_path, sheet_name='角色活动祈愿')
except ValueError:
    print("警告: 无法找到'角色活动祈愿'相关数据")
    print('请检查Excel文档是否损坏或被篡改')
    exit()

df_original_datas_of_character = pd.DataFrame(df_character)

# 将 'time' 列转换为日期时间类型
# errors='coerce' 会将无法解析的日期时间值转换为 NaT
df_original_datas_of_character['time'] = pd.to_datetime(df_original_datas_of_character['time'], errors='coerce')

# === 修改这里：将 NaT 数据设置成其他特殊的值 ===
# 方案1：替换为特定的日期时间 (推荐，保留数据类型语义)
# 例如，替换为Unix纪元（1970-01-01）或者一个你认为合适的默认/未知日期
# 也可以使用 pd.NaT 来保持它是缺失值，但如果你想可视化，那还是需要一个具体值
fill_value_for_time = pd.Timestamp('1900-01-01 00:00:00') # 或者其他你认为合适的日期
df_original_datas_of_character['time'].fillna(fill_value_for_time, inplace=True)

# 方案2：替换为字符串（不推荐，因为会改变列类型，且无法再进行日期时间操作）
# fill_value_for_time_str = '无效日期'
# df_original_datas_of_character['time'].fillna(fill_value_for_time_str, inplace=True)

# 方案3：替换为数字（不推荐，除非你真的想用数字代表日期，这通常不是好做法）
# df_original_datas_of_character['time'].fillna(0, inplace=True) # 这会把日期时间列变成数值列
# ===================================================

# 获取rarity列中为5的行 (五星出金记录)
five_star_items = df_original_datas_of_character[df_original_datas_of_character['rarity'] == 5].copy()

# 确保 'time' 列是字符串类型，因为作为条形图的X轴类别标签时，Pandas的plot方法可能将其转换为索引
# 或者，如果你希望使用日期时间索引，可以设置为索引：
# five_star_items.set_index('time', inplace=True)
# 但对于条形图，通常是类别标签，所以转换为字符串更直接
# 1. 修改X轴时间格式为只显示年月日
five_star_items['time_str'] = five_star_items['time'].dt.strftime('%Y-%m-%d') # 只保留年月日


# 计算平均出金 Pity
mean_pity = five_star_items['within pity'].mean()

# 设置绘图风格
sns.set_style("whitegrid")

fig, ax = plt.subplots(figsize=(15, 8)) # 增大图表尺寸以容纳更多信息

# 绘制柱状图
bars = ax.bar(five_star_items['time_str'], five_star_items['within pity'],
            color='lightcoral', label='5星角色出金Pity')

# --- 添加信息 ---

# 2. 在每个柱子上添加新的列表内容 (例如角色/武器名称)
# 假设你的 five_star_items DataFrame 中有一个名为 'item' 的列，存储了角色/武器名称
item_column_name = 'item' # <--- 请根据你的实际列名修改这里！

for i, bar in enumerate(bars):
    if item_column_name in five_star_items.columns:
        item_name = five_star_items.iloc[i][item_column_name] # 获取对应柱子的物品名称
        yval = bar.get_height()

        ax.text(bar.get_x() + bar.get_width() / 2, # X坐标：柱子中心
                yval + 2, # Y坐标：柱子高度 + 2 (向上偏移)
                str(item_name), # 显示物品名称 (转换为字符串)
                ha='center', va='bottom', # 水平居中，垂直底部对齐
                fontsize=8, color='darkblue', # 字体大小和颜色
                rotation=45) # 旋转文本以防止重叠，尤其是名称较长时
    else:
        print(f"警告：DataFrame中未找到列 '{item_column_name}'，无法在柱子上添加物品名称。")
        break

# 2. 添加平均出金 Pity 线
ax.axhline(mean_pity, color='blue', linestyle='--', linewidth=2, label=f'平均出金Pity: {mean_pity:.2f}')
ax.text(len(five_star_items) - 1, mean_pity + 1, f'平均: {mean_pity:.1f}',
        color='blue', ha='right', va='bottom', fontsize=10) # 在线上方添加文字标签

# 3. 添加软保底和硬保底线
soft_pity_threshold = 74 # 原神通常74-75开始进入软保底
hard_pity_threshold = 90

ax.axhline(soft_pity_threshold, color='green', linestyle=':', linewidth=1.5, label='软保底 (74抽)')
ax.text(len(five_star_items) - 1, soft_pity_threshold + 1, '软保底',
        color='green', ha='right', va='bottom', fontsize=10)

ax.axhline(hard_pity_threshold, color='red', linestyle=':', linewidth=1.5, label='硬保底 (90抽)')
ax.text(len(five_star_items) - 1, hard_pity_threshold + 1, '硬保底',
        color='red', ha='right', va='bottom', fontsize=10)


# --- 设置图表标题和标签 ---
ax.set_title('5星角色出金Pity曲线', fontsize=18) # 更好的标题
ax.set_xlabel('出金日期', fontsize=14) # 修正X轴标签为日期
ax.set_ylabel('抽取Pity', fontsize=14) # 修正Y轴标签

plt.xticks(rotation=60, ha='right', fontsize=10) # 旋转X轴标签，避免重叠
plt.yticks(fontsize=10)

plt.legend(loc='upper left', fontsize=12) # 调整图例位置和大小
plt.tight_layout() # 自动调整布局，防止标签重叠
plt.grid(axis='y', linestyle='--', alpha=0.7) # 添加Y轴网格线

plt.show()

# 此外，你还可以打印一些统计信息到控制台，作为对图表的补充：
print("\n--- 5星角色出金Pity统计信息 ---")
print(f"总计5星出金次数: {len(five_star_items)} 次")
print(f"平均出金Pity: {mean_pity:.2f} 抽")
print(f"最小出金Pity: {five_star_items['within pity'].min()} 抽")
print(f"最大出金Pity: {five_star_items['within pity'].max()} 抽")
print(f"出金Pity中位数: {five_star_items['within pity'].median()} 抽")