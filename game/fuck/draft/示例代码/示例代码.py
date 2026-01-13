import pandas as pd
# import matplotlib.pyplot as plt # 未使用
# import seaborn as sns           # 未使用
# import numpy as np              # 未使用
import os

# 配置excel相关文件
excel_file_path = r'D:\Project\Numpy_Pandas\game\data\Genshin wish logger.xlsx'
file_output_path = r'D:\Project\Numpy_Pandas\game\data\English logger.xlsx'

try:
    # 仅仅是为了验证文件和初始工作表是否存在，实际的数据读取将在后面进行
    # 如果文件不存在，这里也会捕获 FileNotFoundError
    df_character = pd.read_excel(excel_file_path, sheet_name = '角色活动祈愿')
    # df_character_original = pd.DataFrame(df_character) # 冗余，df_character 已经是DataFrame
    # five_character_stars = df_character_original[df_character_original['rarity'] == 5].copy() # 未使用
except ValueError:
    print("警告: 无法找到'角色活动祈愿'相关数据")
    print('请检查Excel文档是否损坏或被篡改')
    exit()
except FileNotFoundError: # 添加对 FileNotFoundError 的捕获
    print(f"警告: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
    exit()

column_to_convert = 'name'

list_map = {
    '琴': 'Jean', '迪卢克': 'Diluc', '温迪': 'Venti', '可莉': 'klee', '莫娜': 'Mona',
    '阿贝多': 'Albedo', '优菈': 'Eula', '魈': 'Xiao', '刻晴': 'Keqing', '七七': 'Qiqi',
    '达达利亚': 'Tartaglia', '钟离': 'Zhongli', '甘雨': 'Ganyu', '胡桃': 'Hu Tao',
    '申鹤': 'Shenhe', '夜兰': 'Yelan', '白术': 'Baizhu', '闲云': 'Xianyun',
    '神里绫华': 'Kamisato Ayaka', '枫原万叶': 'Kaedehara Kazuha', '宵宫': 'Yoimiya',
    '雷电将军': 'Raiden Shogun', '珊瑚宫心海': 'Sangonomiya Kokomi', '荒泷一斗': 'Arataki Itto',
    '八重神子': 'Yae Miko', '神里绫人': 'Kamisato Ayato', '梦见月瑞希': 'Yumemizuki Mizuki',
    '提纳里': 'Tighnari', '赛诺': 'Cyno', '妮露': 'Nilou', '纳西妲': 'Nahida',
    '流浪者': 'Wanderer', '艾尔海森': 'Alhaitham', '迪希雅': 'Dehya', '林尼': 'Lyney',
    '那维莱特': 'Neuvillette', '莱欧斯利': 'Wirithesley', '芙宁娜': 'Furina',
    '纳维娅': 'Navia', '千织': 'Chiori', '阿蕾奇诺': 'Arlecchino', '克洛琳德': 'Clorinde',
    '希格雯': 'Sigewinne', '艾梅莉埃': 'Emilie', '爱可菲': 'Escoffier', '玛拉妮': 'Mualani',
    '基尼奇': 'Kinich', '希诺宁': 'Xilonen', '恰斯卡': 'Chasca', '茜特拉莉': 'Citlali',
    '玛薇卡': 'Mavuika', '瓦雷莎': 'Varesa', '伊法': 'Ifa', '丝柯克': 'Skirk',
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
    # --- 1. 读取所有Excel工作表数据 ---
    # 修正 1: pd.read_excel 的第一个参数应该是文件路径
    all_sheet_data = pd.read_excel(excel_file_path, sheet_name = None)
    # 解释:
    # - pd.read_excel(): 这是Pandas库中用于读取Excel文件的函数。
    # - excel_file_path: 这是一个字符串变量，存储着你要读取的原始Excel文件的完整路径。
    # - sheet_name = None: 这是read_excel函数的一个关键参数。当设置为None时，Pandas会读取Excel文件中所有可见的工作表。
    # - all_sheet_data: 这是一个变量，用于存储read_excel函数的返回值。当sheet_name=None时，返回值是一个OrderedDict（有序字典）。
    #   - 字典的键（keys）是Excel文件中各个工作表的原始名称（例如，'角色活动祈愿', '武器活动祈愿'）。
    #   - 字典的值（values）是每个工作表对应的数据，以Pandas DataFrame（数据框）的形式存储。

    converted_sheets = {}
    # 解释:
    # - converted_sheets: 这是一个空字典，用于在接下来的循环中存储处理后的（即可能已更改名称和/或内容）的DataFrame。
    #   - 它的键将是新的（英文或未变动的）工作表名称。
    #   - 它的值将是对应的DataFrame。

    # --- 2. 遍历并处理每个工作表 ---
    for original_sheet_name, df in all_sheet_data.items():
        # 解释:
        # - for original_sheet_name, df in all_sheet_data.items(): 这是一个Python的for循环，用于遍历all_sheet_data字典中的所有键值对。
        #   - all_sheet_data.items(): 返回字典中所有键值对的视图。
        #   - original_sheet_name: 在每次循环中，这个变量会接收当前工作表的原始（中文）名称，例如 '角色活动祈愿'。
        #   - df: 在每次循环中，这个变量会接收当前工作表对应的Pandas DataFrame。

        new_sheet_name = sheet_map.get(original_sheet_name, original_sheet_name)
        # 解释:
        # - sheet_map: 这是一个预先定义的字典变量，它存储了原始中文工作表名称到英文名称的映射关系（例如 {'角色活动祈愿': 'Character Event Prayers', ...}）。
        # - .get(key, default_value): 这是字典的一个方法。它尝试从sheet_map字典中获取original_sheet_name对应的值。
        #   - 如果original_sheet_name在sheet_map中找到了对应的英文名称，那么new_sheet_name就会被赋值为那个英文名称。
        #   - 如果original_sheet_name在sheet_map中没有找到（即没有预设的翻译），那么new_sheet_name就会被赋值为default_value，也就是original_sheet_name本身。这意味着该工作表的名称将保持不变。
        # - new_sheet_name: 这个变量存储了当前工作表的新名称（可能是翻译后的英文名，也可能是保持不变的原始名）。

        # 总是对数据框进行操作，然后将其添加到 converted_sheets
        if column_to_convert in df.columns:
            # 解释:
            # - if column_to_convert in df.columns: 这是一个条件语句。
            #   - column_to_convert: 这是一个字符串变量（例如 'name'），表示要进行内容转换的列名。
            #   - df.columns: 这是DataFrame df的一个属性，它返回DataFrame中所有列名的集合。
            #   - `in` 运算符：检查 `column_to_convert` 是否存在于当前DataFrame `df` 的列名中。
            #   - 如果当前工作表包含名为 `column_to_convert` 的列（例如 'name' 列），则执行if块内的代码。

            df[column_to_convert] = df[column_to_convert].replace(list_map)
            # 解释:
            # - df[column_to_convert]: 选中DataFrame `df` 中 `column_to_convert` 指定的列（例如 'name' 列）。这会返回一个Pandas Series（类似于一维数组）。
            # - .replace(list_map): 这是Pandas Series的一个方法，用于替换Series中的值。
            #   - list_map: 这是一个预先定义的字典变量，它存储了中文角色名到英文角色名的映射关系（例如 {'琴': 'Jean', '迪卢克': 'Diluc', ...}）。
            #   - 这个方法会遍历 `df[column_to_convert]` 这一列的每一个值。如果某个值是 `list_map` 字典中的一个键，那么它就会被替换为 `list_map` 中对应的值。
            #   - `=` 赋值操作：将替换后的Series重新赋值回 `df` 的 `column_to_convert` 列，从而更新了DataFrame中的数据。

        else:
            print(f"Warning: Column '{column_to_convert}' does not exist in Sheet '{original_sheet_name}', skipping column data conversion.")
            # 解释:
            # - 如果 `if` 条件（即 `column_to_convert` 列不存在）不满足，则执行此 `else` 块。
            # - print(...): 打印一个警告信息，说明在当前工作表（由 `original_sheet_name` 指定）中没有找到需要转换的列，因此跳过了该列的数据转换。
            # - f-string (f"...{}"): 一种Python格式化字符串的方式，允许直接在字符串中嵌入变量（如 `{column_to_convert}` 和 `{original_sheet_name}`）。

        # 修正 2: 这行代码必须在 for 循环内部，确保每个 DataFrame 都被添加
        converted_sheets[new_sheet_name] = df
        # 解释:
        # - 这行代码将当前处理过的DataFrame `df` (无论其内容是否被转换，或者列名是否被检查过) 添加到 `converted_sheets` 字典中。
        # - new_sheet_name: 作为字典的键，即该DataFrame在最终Excel文件中将拥有的工作表名称。
        # - df: 作为字典的值，即该工作表对应的数据DataFrame。
        # - 这一步对于确保所有原始工作表（无论是否需要名称转换或内容转换）都被收集起来以便后续写入新文件至关重要。

    # --- 3. 将处理后的数据写入新的Excel文件 ---
    with pd.ExcelWriter(file_output_path, engine = 'openpyxl') as writer:
        # 解释:
        # - with ... as ...: 这是一个Python的上下文管理器。它确保资源（如文件）在使用完毕后能被正确关闭，即使发生错误。
        # - pd.ExcelWriter(): 这是Pandas中用于创建和管理Excel文件的写入器对象。
        # - file_output_path: 这是一个字符串变量，存储着将要创建的新Excel文件的完整路径。
        # - engine = 'openpyxl': 指定Pandas使用哪个后端库来处理Excel文件。'openpyxl'是处理.xlsx格式文件的常用且推荐的引擎。
        # - as writer: 将创建的Excel写入器对象赋值给变量 `writer`，我们将在后续步骤中使用它来写入数据。

        for sheet_name, df_to_writer in converted_sheets.items():
            # 解释:
            # - 这个for循环遍历 `converted_sheets` 字典，该字典包含了所有准备写入新Excel文件的DataFrame及其对应的（新）工作表名称。
            # - sheet_name: 在每次循环中，这个变量会接收当前要写入的工作表的名称（例如 'Character Event Prayers'）。
            # - df_to_writer: 在每次循环中，这个变量会接收当前要写入的DataFrame。

            df_to_writer.to_excel(writer, sheet_name = sheet_name, index = False)
            # 解释:
            # - df_to_writer.to_excel(): 这是DataFrame的一个方法，用于将DataFrame写入Excel文件。
            # - writer: 指定使用之前创建的 `ExcelWriter` 对象来写入数据。
            # - sheet_name = sheet_name: 指定在新的Excel文件中，当前DataFrame将被写入到哪个工作表名称下。
            # - index = False: 这是一个重要参数。默认情况下，Pandas会将DataFrame的行索引（即0, 1, 2...）作为Excel文件的第一列写入。设置 `index=False` 可以防止写入这些索引，使输出的Excel表格更整洁，只包含你原始的数据列。

    print(f"Excel文件已成功转换为英文并保存到: {file_output_path}")
    # 解释:
    # - 当所有操作都成功完成（没有抛出任何异常）时，这行代码会被执行，打印一条成功消息，并告知用户新文件的保存路径。

# --- 4. 异常处理 ---
except FileNotFoundError:
    print(f"错误: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
    # 解释:
    # - except FileNotFoundError: 捕获 `FileNotFoundError` 异常。如果 `pd.read_excel` 或 `pd.ExcelWriter` 在尝试访问文件时发现文件不存在，就会抛出此异常。
    # - print(...): 打印一个错误消息，通知用户文件未找到，并建议检查路径。

except KeyError as e:
    # 更详细的KeyError提示
    print(f"数据处理错误: 无法找到预期的键/列或映射失败: {e}. 请检查Excel文件的数据结构或映射字典。")
    # 解释:
    # - except KeyError as e: 捕获 `KeyError` 异常。这种异常通常发生在尝试访问字典中不存在的键（例如 `list_map` 或 `sheet_map` 中没有对应的映射）或DataFrame中不存在的列（但此代码中 `if column_to_convert in df.columns:` 已经做了大部分处理）。`as e` 会将具体的错误信息捕获到变量 `e` 中。
    # - print(...): 打印一个错误消息，更详细地说明可能是键或列查找失败，并建议用户检查数据结构或映射字典。

except Exception as e:
    print(f"发生未知错误: {e}")
    # 解释:
    # - except Exception as e: 这是一个通用的异常捕获块。如果发生上述特定异常之外的任何其他错误，都将被这个块捕获。
    # - 强烈建议将 `Exception` 放在所有其他更具体的 `except` 块之后，因为它会捕获所有类型的异常，这可能会掩盖更具体的错误信息。
    # - print(...): 打印一个通用的错误消息，指出发生了未知错误，并显示具体的错误信息 `e`。

except FileNotFoundError:
    print(f"错误: 文件 '{excel_file_path}' 未找到。请检查文件路径和名称是否正确。")
except KeyError as e:
    # 更详细的KeyError提示
    print(f"数据处理错误: 无法找到预期的键/列或映射失败: {e}. 请检查Excel文件的数据结构或映射字典。")
except Exception as e:
    print(f"发生未知错误: {e}")