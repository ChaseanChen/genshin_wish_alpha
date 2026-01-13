import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np # 用于生成示例数据中的NaN

# --- 0. 准备示例数据 (如果你没有Excel文件，可以运行此部分来生成一个模拟文件) ---
# 这个模拟文件将包含 '日期', '产品类别', '销售额', '成本', '地区' 等列
try:
    # 尝试读取，如果文件不存在，则创建模拟数据
    df = pd.read_excel('你的数据.xlsx', sheet_name='Sheet1')
    print("已找到并读取 '你的数据.xlsx'。")
except FileNotFoundError:
    print("未找到 '你的数据.xlsx'，正在生成模拟数据文件。")
    # 生成模拟数据
    data = {
        '日期': pd.to_datetime(['2023-01-15', '2023-01-20', '2023-02-01', '2023-02-10', '2023-03-05',
                               '2023-03-12', '2023-04-01', '2023-04-20', '2023-05-01', '2023-05-15',
                               '2023-01-25', '2023-02-28', '2023-03-30', '2023-04-10', '2023-05-20']),
        '产品类别': ['电子产品', '服装', '食品', '电子产品', '服装',
                     '食品', '电子产品', '服装', '食品', '电子产品',
                     '电子产品', '服装', '食品', '电子产品', '服装'],
        '销售额': [1200, 300, 150, 900, 450,
                   200, 1500, 500, 250, 1100,
                   1000, 350, 180, 800, 400],
        '成本': [800, 150, 80, 600, 200,
                 100, 1000, 250, 120, 700,
                 650, 180, 90, 500, 220],
        '地区': ['华东', '华北', '华南', '华东', '华北',
                 '华南', '华东', '华北', '华南', '华东',
                 '华北', '华南', '华东', '华北', '华南'],
        '客户ID': ['C001', 'C002', 'C003', 'C004', 'C005',
                   'C006', 'C007', 'C008', 'C009', 'C010',
                   'C011', 'C012', 'C013', 'C014', 'C015']
    }
    # 增加一些缺失值和重复值来模拟真实数据
    data['销售额'][2] = np.nan # 模拟缺失值
    data['产品类别'][5] = '服装' # 模拟重复值
    data['日期'][1] = pd.NaT # 模拟无效日期
    data['销售额'] = data['销售额'] * 1.0 # 确保是浮点数
    
    df = pd.DataFrame(data)
    df.to_excel('你的数据.xlsx', index=False, sheet_name='Sheet1')
    print("模拟数据文件 '你的数据.xlsx' 已创建。请运行此脚本。")

# --- 设置 Matplotlib 支持中文显示 ---
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# --- 1. 读取Excel文档内容 ---
file_path = '你的数据.xlsx'
sheet_name = 'Sheet1'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print("\n--- Excel文件读取成功 ---")
    print("数据前5行：\n", df.head())
    print("\n数据基本信息：")
    df.info()
except FileNotFoundError:
    print(f"错误：文件 '{file_path}' 未找到。请检查文件路径和文件名。")
    exit() # 如果文件没找到，就退出程序
except Exception as e:
    print(f"读取Excel时发生错误：{e}")
    exit()

# --- 2. 数据初步探索与清洗 ---
print("\n--- 数据清洗与预处理 ---")

# 1. 处理缺失值 (删除含有缺失值的行，简单示例)
print("\n每列的缺失值数量：\n", df.isnull().sum())
df_processed = df.dropna()
print(f"\n删除缺失值行后，剩余行数：{len(df_processed)}")

# 2. 处理重复值
print("\n重复行数量：", df_processed.duplicated().sum())
df_processed.drop_duplicates(inplace=True)
print(f"删除重复行后，剩余行数：{len(df_processed)}")

# 3. 数据类型转换 (确保日期是datetime，数值是numeric)
if '日期' in df_processed.columns:
    df_processed['日期'] = pd.to_datetime(df_processed['日期'], errors='coerce')
    df_processed.dropna(subset=['日期'], inplace=True) # 删除无法转换为日期的行
if '销售额' in df_processed.columns:
    df_processed['销售额'] = pd.to_numeric(df_processed['销售额'], errors='coerce')
    df_processed.dropna(subset=['销售额'], inplace=True)
if '成本' in df_processed.columns:
    df_processed['成本'] = pd.to_numeric(df_processed['成本'], errors='coerce')
    df_processed.dropna(subset=['成本'], inplace=True)

print("\n处理后的数据基本信息：")
df_processed.info()
print("\n处理后的数据前5行：\n", df_processed.head())


# --- 3. 数据处理与分析 ---
print("\n--- 数据处理与分析 ---")

# 计算总利润
if '销售额' in df_processed.columns and '成本' in df_processed.columns:
    df_processed['利润'] = df_processed['销售额'] - df_processed['成本']
    total_profit = df_processed['利润'].sum()
    print(f"\n总利润：{total_profit:.2f}")
else:
    print("\n'销售额' 或 '成本' 列不存在，无法计算利润。")

# 各产品类别的总销售额
sales_by_category = None
if '产品类别' in df_processed.columns and '销售额' in df_processed.columns:
    sales_by_category = df_processed.groupby('产品类别')['销售额'].sum().sort_values(ascending=False)
    print("\n各产品类别的总销售额：\n", sales_by_category)
else:
    print("\n'产品类别' 或 '销售额' 列不存在，无法按类别分析销售额。")

# 每月销售额趋势
sales_by_month = None
if '日期' in df_processed.columns and '销售额' in df_processed.columns:
    df_processed['月份'] = df_processed['日期'].dt.to_period('M')
    sales_by_month = df_processed.groupby('月份')['销售额'].sum().sort_index()
    print("\n每月总销售额：\n", sales_by_month)
else:
    print("\n'日期' 或 '销售额' 列不存在，无法分析销售趋势。")

# 各地区的平均利润
avg_profit_by_region = None
if '地区' in df_processed.columns and '利润' in df_processed.columns:
    avg_profit_by_region = df_processed.groupby('地区')['利润'].mean().sort_values(ascending=False)
    print("\n各地区的平均利润：\n", avg_profit_by_region)
else:
    print("\n'地区' 或 '利润' 列不存在，无法分析地区平均利润。")


# --- 4. 数据可视化 ---
print("\n--- 数据可视化 ---")

# 1. 柱状图：各产品类别总销售额
if sales_by_category is not None and not sales_by_category.empty:
    plt.figure(figsize=(10, 6))
    sns.barplot(x=sales_by_category.index, y=sales_by_category.values, palette='viridis')
    plt.title('各产品类别总销售额')
    plt.xlabel('产品类别')
    plt.ylabel('总销售额')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('各产品类别总销售额.png', dpi=300)
    plt.show()
else:
    print("\n跳过柱状图绘制：'销售额按产品类别汇总数据' 不可用或为空。")

# 2. 折线图：每月销售额趋势
if sales_by_month is not None and not sales_by_month.empty:
    plt.figure(figsize=(12, 6))
    sales_by_month.index = sales_by_month.index.astype(str) # 将PeriodIndex转换为字符串
    sns.lineplot(x=sales_by_month.index, y=sales_by_month.values, marker='o', color='blue')
    plt.title('每月销售额趋势')
    plt.xlabel('月份')
    plt.ylabel('总销售额')
    plt.xticks(rotation=45, ha='right')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig('每月销售额趋势.png', dpi=300)
    plt.show()
else:
    print("\n跳过折线图绘制：'每月销售额趋势数据' 不可用或为空。")

# 3. 散点图：销售额与利润的关系
if '销售额' in df_processed.columns and '利润' in df_processed.columns and not df_processed.empty:
    plt.figure(figsize=(8, 6))
    sns.scatterplot(x='销售额', y='利润', data=df_processed, 
                    hue='产品类别' if '产品类别' in df_processed.columns else None, 
                    palette='coolwarm', alpha=0.7)
    plt.title('销售额与利润的关系')
    plt.xlabel('销售额')
    plt.ylabel('利润')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig('销售额与利润的关系.png', dpi=300)
    plt.show()
else:
    print("\n跳过散点图绘制：'销售额' 或 '利润' 列不存在，或数据为空。")

# 4. 饼图：各地区销售额占比
if '地区' in df_processed.columns and '销售额' in df_processed.columns and not df_processed.empty:
    sales_by_region = df_processed.groupby('地区')['销售额'].sum()
    if not sales_by_region.empty:
        plt.figure(figsize=(8, 8))
        plt.pie(sales_by_region, labels=sales_by_region.index, autopct='%1.1f%%', startangle=90, colors=sns.color_palette('pastel'))
        plt.title('各地区销售额占比')
        plt.axis('equal') # 保证饼图是圆形
        plt.tight_layout()
        plt.savefig('各地区销售额占比.png', dpi=300)
        plt.show()
    else:
        print("\n跳过饼图绘制：'销售额按地区汇总数据' 为空。")
else:
    print("\n跳过饼图绘制：'地区' 或 '销售额' 列不存在，或数据为空。")

print("\n--- 所有操作完成 ---")