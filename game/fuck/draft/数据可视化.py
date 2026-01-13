import matplotlib.pyplot as plt
import seaborn as sns

# 设置 Matplotlib 支持中文显示
plt.rcParams['font.sans-serif'] = ['SimHei'] # 指定默认字体为黑体
plt.rcParams['axes.unicode_minus'] = False # 解决保存图像时负号'-'显示为方块的问题

print("\n--- 数据可视化 ---")

# 1. 柱状图：各产品类别的总销售额
if '产品类别' in df_processed.columns and '销售额' in df_processed.columns:
    plt.figure(figsize=(10, 6))
    sales_by_category = df_processed.groupby('产品类别')['销售额'].sum().sort_values(ascending=False)
    sns.barplot(x=sales_by_category.index, y=sales_by_category.values, palette='viridis')
    plt.title('各产品类别总销售额')
    plt.xlabel('产品类别')
    plt.ylabel('总销售额')
    plt.xticks(rotation=45, ha='right') # 旋转x轴标签，防止重叠
    plt.tight_layout() # 自动调整子图参数，使之填充整个图像区域
    plt.show()
else:
    print("\n跳过柱状图：'产品类别' 或 '销售额' 列不存在。")

# 2. 折线图：每月销售额趋势
if '日期' in df_processed.columns and '销售额' in df_processed.columns:
    # 确保 '日期' 列已转换为 datetime 类型
    df_processed['日期'] = pd.to_datetime(df_processed['日期'], errors='coerce')
    df_processed_date_valid = df_processed.dropna(subset=['日期'])
    df_processed_date_valid['月份'] = df_processed_date_valid['日期'].dt.to_period('M')
    sales_by_month = df_processed_date_valid.groupby('月份')['销售额'].sum().sort_index()

    plt.figure(figsize=(12, 6))
    # 将 PeriodIndex 转换为字符串以便于绘图，或者直接转换为时间戳
    sales_by_month.index = sales_by_month.index.astype(str)
    sns.lineplot(x=sales_by_month.index, y=sales_by_month.values, marker='o', color='blue')
    plt.title('每月销售额趋势')
    plt.xlabel('月份')
    plt.ylabel('总销售额')
    plt.xticks(rotation=45, ha='right')
    plt.grid(True, linestyle='--', alpha=0.6) # 添加网格线
    plt.tight_layout()
    plt.show()
else:
    print("\n跳过折线图：'日期' 或 '销售额' 列不存在。")

# 3. 散点图：销售额与利润的关系 (如果存在 '销售额' 和 '利润' 列)
if '销售额' in df_processed.columns and '利润' in df_processed.columns:
    plt.figure(figsize=(8, 6))
    sns.scatterplot(x='销售额', y='利润', data=df_processed, hue='产品类别' if '产品类别' in df_processed.columns else None, palette='coolwarm', alpha=0.7)
    plt.title('销售额与利润的关系')
    plt.xlabel('销售额')
    plt.ylabel('利润')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.show()
else:
    print("\n跳过散点图：'销售额' 或 '利润' 列不存在。")

# 4. 饼图：各地区销售额占比 (如果存在 '地区' 和 '销售额' 列)
if '地区' in df_processed.columns and '销售额' in df_processed.columns:
    sales_by_region = df_processed.groupby('地区')['销售额'].sum()
    plt.figure(figsize=(8, 8))
    plt.pie(sales_by_region, labels=sales_by_region.index, autopct='%1.1f%%', startangle=90, colors=sns.color_palette('pastel'))
    plt.title('各地区销售额占比')
    plt.axis('equal') # 保证饼图是圆形
    plt.tight_layout()
    plt.show()
else:
    print("\n跳过饼图：'地区' 或 '销售额' 列不存在。")

# 你可以根据你的数据和想表达的洞察，绘制更多类型的图表，例如：
# - 直方图：某个数值列的分布 (如：销售额的分布)
# - 箱线图：不同类别下某个数值列的分布 (如：不同产品类别的销售额分布)
# - 热力图：两个分类变量之间的交叉关系