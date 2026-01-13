# 在上面每个 plt.show() 之前，添加保存代码

# 1. 柱状图：各产品类别的总销售额
if '产品类别' in df_processed.columns and '销售额' in df_processed.columns:
    plt.figure(figsize=(10, 6))
    sales_by_category = df_processed.groupby('产品类别')['销售额'].sum().sort_values(ascending=False)
    sns.barplot(x=sales_by_category.index, y=sales_by_category.values, palette='viridis')
    plt.title('各产品类别总销售额')
    plt.xlabel('产品类别')
    plt.ylabel('总销售额')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('各产品类别总销售额.png', dpi=300) # 保存图片，dpi表示分辨率
    plt.show()

# 2. 折线图：每月销售额趋势
if '日期' in df_processed.columns and '销售额' in df_processed.columns:
    # ... (省略计算 sales_by_month 的代码，与上面相同)
    df_processed['日期'] = pd.to_datetime(df_processed['日期'], errors='coerce')
    df_processed_date_valid = df_processed.dropna(subset=['日期'])
    df_processed_date_valid['月份'] = df_processed_date_valid['日期'].dt.to_period('M')
    sales_by_month = df_processed_date_valid.groupby('月份')['销售额'].sum().sort_index()
    sales_by_month.index = sales_by_month.index.astype(str)

    plt.figure(figsize=(12, 6))
    sns.lineplot(x=sales_by_month.index, y=sales_by_month.values, marker='o', color='blue')
    plt.title('每月销售额趋势')
    plt.xlabel('月份')
    plt.ylabel('总销售额')
    plt.xticks(rotation=45, ha='right')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig('每月销售额趋势.png', dpi=300)
    plt.show()

# ... 其他图表的保存方法类似