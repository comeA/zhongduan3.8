#专门处理 筛选过后的 的 业务号码 数据复制到 导入模板文件

import pandas as pd
from modules.excel_utils import copy_data_to_excel

def copy_business_numbers_to_template(df, template_filepath):
    """将筛选后的“业务号码”复制到模板文件"""
    try:
        business_numbers = df['业务号码'].astype(str).str.lstrip("'")
        if not copy_data_to_excel(business_numbers, template_filepath, "Sheet1", "业务号码"):
            print("复制业务号码失败")
            return False
        print("业务号码已成功复制到 导入模板.xlsx")
        return True
    except KeyError:
        print("错误：筛选后的数据中缺少 '业务号码' 列。")
        return False
    except Exception as e:
        print(f"处理业务号码或写入“导入模板.xlsx”文件时发生错误：{e}")
        return False