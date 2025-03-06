
import pandas as pd

def truncate_sheet_name(sheet_name, max_length=31):
    """
    截断工作表名称，确保其长度不超过 max_length 个字符。
    :param sheet_name: 原始工作表名称
    :param max_length: 最大允许长度（默认为 31）
    :return: 截断后的工作表名称
    """
    if len(sheet_name) > max_length:
        return sheet_name[:max_length]
    return sheet_name


# 定义 process_data 函数
def process_data(filtered_sheet_name, filtered_filepath, df):
    # 调用 truncate_sheet_name
    updated_sheet_name = truncate_sheet_name(filtered_sheet_name + "_更新精简型号")

    # 保存更新后的数据到目标文件
    with pd.ExcelWriter(filtered_filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=updated_sheet_name, index=False)
    print(f"数据已保存到 {filtered_filepath} 的 {updated_sheet_name} 工作表中。")