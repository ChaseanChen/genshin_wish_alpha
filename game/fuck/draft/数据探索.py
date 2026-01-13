# import pandas as pd

# 假设我们已经成功读取了 df

print("\n--- 数据清洗与预处理 ---")

# 1. 查看缺失值
print("\n每列的缺失值数量：")
print(df.isnull().sum())

# 示例：处理缺失值 (选择一种或多种策略)
# a. 删除含有缺失值的行 (简单粗暴，可能丢失有用数据)
# df_cleaned = df.dropna()
# b. 填充缺失值 (例如，用0、平均值、中位数或众数填充)
# df['某一列'].fillna(0, inplace=True) # 用0填充
# df['某一列'].fillna(df['某一列'].mean(), inplace=True) # 用平均值填充
# 这里我们选择一个简单的策略：如果你的数据量大，可以删除少量缺失值行
df_processed = df.dropna()
print(f"\n删除缺失值行后，剩余行数：{len(df_processed)}")


# 2. 查看重复值
print("\n重复行数量：", df_processed.duplicated().sum())
# 删除重复值
df_processed.drop_duplicates(inplace=True)
print(f"删除重复行后，剩余行数：{len(df_processed)}")

# 3. 数据类型转换 (示例：假设有一列名为'日期'，需要转换为日期时间类型)
# 如果你的Excel里有日期列，并且 Pandas 没有正确识别，需要手动转换
# df_processed['日期'] = pd.to_datetime(df_processed['日期'], errors='coerce')
# errors='coerce' 会把无法转换的值设为 NaT (Not a Time)，可以后续处理

# 4. 列名标准化 (如果列名有空格、特殊字符或不规范)
# df_processed.columns = df_processed.columns.str.strip().str.replace(' ', '_').str.lower()
# print("\n标准化后的列名：", df_processed.columns.tolist())

# 再次查看处理后的数据信息
print("\n处理后的数据基本信息：")
df_processed.info()
print("\n处理后的数据前5行：")
print(df_processed.head())