# 假设 df_processed 是经过清洗后的DataFrame

print("\n--- 数据处理与分析 ---")

# 示例1：计算每个产品类别的总销售额
if '产品类别' in df_processed.columns and '销售额' in df_processed.columns:
    sales_by_category = df_processed.groupby('产品类别')['销售额'].sum().sort_values(ascending=False)
    print("\n各产品类别的总销售额：")
    print(sales_by_category)
else:
    print("\n'产品类别' 或 '销售额' 列不存在，跳过此分析。请根据实际列名修改。")

# 示例2：计算总利润 (需要有 '销售额' 和 '成本' 列)
if '销售额' in df_processed.columns and '成本' in df_processed.columns:
    df_processed['利润'] = df_processed['销售额'] - df_processed['成本']
    total_profit = df_processed['利润'].sum()
    print(f"\n总利润：{total_profit:.2f}")

    # 示例3：每个地区的平均利润 (需要 '地区' 和 '利润' 列)
    if '地区' in df_processed.columns:
        avg_profit_by_region = df_processed.groupby('地区')['利润'].mean().sort_values(ascending=False)
        print("\n各地区的平均利润：")
        print(avg_profit_by_region)
    else:
        print("\n'地区' 列不存在，跳过地区平均利润分析。")
else:
    print("\n'销售额' 或 '成本' 列不存在，跳过利润相关分析。请根据实际列名修改。")


# 示例4：按日期查看销售趋势 (需要有 '日期' 和 '销售额' 列)
# 确保 '日期' 列是 datetime 类型
if '日期' in df_processed.columns and '销售额' in df_processed.columns:
    df_processed['日期'] = pd.to_datetime(df_processed['日期'], errors='coerce')
    # 删除日期为 NaT 的行 (如果有)
    df_processed_date_valid = df_processed.dropna(subset=['日期'])
    
    # 按月聚合销售额
    df_processed_date_valid['月份'] = df_processed_date_valid['日期'].dt.to_period('M')
    sales_by_month = df_processed_date_valid.groupby('月份')['销售额'].sum().sort_index()
    print("\n每月总销售额：")
    print(sales_by_month)
else:
    print("\n'日期' 或 '销售额' 列不存在，跳过销售趋势分析。请根据实际列名修改。")