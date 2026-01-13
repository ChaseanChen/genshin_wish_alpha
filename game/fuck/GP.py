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

# === 在这里添加转换代码 ===
# 将 'time' 列转换为日期时间类型
# errors='coerce' 会将无法解析的日期时间值转换为 NaT (Not a Time)
df_original_datas_of_character['time'] = pd.to_datetime(df_original_datas_of_character['time'], errors='coerce')

# 移除因转换失败（例如，原始数据有空值或错误格式）而产生的 NaT 行（可选，但推荐）
df_original_datas_of_character.dropna(subset=['time'], inplace=True)
# === 转换代码结束 ===


# 获取rarity列中为5的行 (五星出金记录)
five_star_items = df_original_datas_of_character[df_original_datas_of_character['rarity'] == 5].copy()

# 确保 'time' 列是字符串类型，因为作为条形图的X轴类别标签时，Pandas的plot方法可能将其转换为索引
# 或者，如果你希望使用日期时间索引，可以设置为索引：
# five_star_items.set_index('time', inplace=True)
# 但对于条形图，通常是类别标签，所以转换为字符串更直接
# five_star_items['time_str'] = five_star_items['time'].dt.strftime('%Y-%m-%d %H:%M')
five_star_items['time_str'] = five_star_items['time'].dt.strftime('%Y-%m-%d')

# 计算平均出金 Pity
mean_pity = five_star_items['within pity'].mean()

# 设置绘图风格
sns.set_style("whitegrid")

fig, ax = plt.subplots(figsize=(15, 8)) # 增大图表尺寸以容纳更多信息

# 绘制柱状图
bars = ax.bar(five_star_items['time_str'], five_star_items['within pity'],
        color='lightcoral', label='5-star character cash-out Pity')

# --- 添加信息 ---

# 1. 在每个柱子上添加 Pity 数值标签
for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval + 0.5, # 调整 yval + 0.5 向上偏移一点
                round(yval), ha='center', va='bottom', fontsize=9, color='black') # round(yval) 可以确保显示整数

# 2. 添加平均出金 Pity 线
ax.axhline(mean_pity, color='blue', linestyle='--', linewidth=2, label=f'Average Withdrawal Pity: {mean_pity:.2f}')
ax.text(len(five_star_items) - 1, mean_pity + 1, f'Average: {mean_pity:.1f}',
        color='blue', ha='right', va='bottom', fontsize=10) # 在线上方添加文字标签

# 3. 添加软保底和硬保底线
soft_pity_threshold = 74 # 原神通常74-75开始进入软保底
hard_pity_threshold = 90

ax.axhline(soft_pity_threshold, color='green', linestyle=':', linewidth=1.5, label='Soft Guarantee (74 draws)')
ax.text(len(five_star_items) - 1, soft_pity_threshold + 1, 'Soft Guarantee',
        color='green', ha='right', va='bottom', fontsize=10)

ax.axhline(hard_pity_threshold, color='red', linestyle=':', linewidth=1.5, label='Hard Guarantee (90 draws)')
ax.text(len(five_star_items) - 1, hard_pity_threshold + 1, 'Hard Guarantee',
        color='red', ha='right', va='bottom', fontsize=10)


# --- 设置图表标题和标签 ---
ax.set_title('5-star character cash-out Pity', fontsize=18) # 更好的标题
ax.set_xlabel('time', fontsize=14) # 修正X轴标签
ax.set_ylabel('whthin Pity', fontsize=14) # 修正Y轴标签

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