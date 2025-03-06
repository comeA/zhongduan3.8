'''
程序名称：终端出库报表
版本 ： V4.0
功能： 处理日常的终端数据，包括复制指定数据清单，筛选指定数据，插入指定字段，匹配指定数据
开发人员： 卢鹤斌
优化人员： 卢鹤斌
注意事项 ： 待处理的文件后缀需为 .xlsx 格式 !!!!
主程序 入口 ： main.py 文件

软件使用说明：
   1、该程序所有用到的文件后缀 ， 最好为  .xlsx 格式 ，如遇 .csv格式 请转为 .xlsx格式
   2、源文件和目标文件路径 问题： 先输入“文件路径” ，在输入 “文件名称” ， 最后输入“文件子表名称”（目标文件的子表名称可以自定义）
      dwd_hzluheb_acc_sn_final_pg 文件 问题 ： 先输入“文件路径\文件名称” ，最后在输入“文件子表名称”
   3、除了 第二点 说到 的路径是分步骤输入外，其他文件的输入请直接 输入 文件路径\文件名称

'''

import pandas as pd
import re
# import modules.truncate_sheet_name as process_data


# 定义 truncate_sheet_name 函数
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


def perform_vlookup_correct(df, sn_df):
    """执行正确的 VLOOKUP 操作，考虑数据类型转换和重复键问题"""
    try:
        # 强制转换为字符串类型
        df['业务号码'] = df['业务号码'].astype(str)
        sn_df['rms_access_code'] = sn_df['rms_access_code'].astype(str)
        sn_df['ce_loid'] = sn_df['ce_loid'].astype(str)
        sn_df['create_date'] = pd.to_datetime(sn_df['create_date'])

        # 对SN数据按照rms_access_code分组，并取create_date最大的那一行
        sn_df = sn_df.sort_values(by=['rms_access_code', 'create_date'], ascending=[True, False]).groupby('rms_access_code').head(1)

        # 执行 VLOOKUP
        df = pd.merge(df, sn_df[['rms_access_code', 'ce_loid']], left_on='业务号码', right_on='rms_access_code', how='left')

        # 将匹配到的 LOID（SN码） 赋值给新列
        df['LOID（SN码）'] = df['ce_loid']
        df.drop(columns=['ce_loid','rms_access_code'], inplace=True)

        return df
    except KeyError as e:
        print(f"KeyError: 缺少列：{e}")
        return None
    except Exception as e:
        print(f"VLOOKUP 过程中发生错误：{e}")
        return None





def perform_mac_status_vlookup(df, mac_export_filepath):
    """根据ISCM终端MAC地址字段匹配cpeExport_mac20250303.xlsx表中的终端唯一标识，并填充终端注册状态"""
    try:
        # 读取cpeExport_mac20250303.xlsx文件
        mac_df = pd.read_excel(mac_export_filepath, sheet_name="sheet", engine='openpyxl')

        # 强制转换为字符串类型
        df['ISCM终端MAC地址'] = df['ISCM终端MAC地址'].astype(str)
        mac_df['终端唯一标识'] = mac_df['终端唯一标识'].astype(str)

        # 执行VLOOKUP操作
        df = pd.merge(df, mac_df[['终端唯一标识', '终端注册状态']], left_on='ISCM终端MAC地址', right_on='终端唯一标识', how='left')

        # 将匹配到的终端注册状态赋值给新列
        df['ISCM终端MAC地址-注册状态'] = df['终端注册状态']
        df.drop(columns=['终端唯一标识', '终端注册状态'], inplace=True)

        return df
    except KeyError as e:
        print(f"KeyError: 缺少列：{e}")
        return None
    except Exception as e:
        print(f"VLOOKUP 过程中发生错误：{e}")
        return None

#
# def perform_simplified_model_vlookup(df, simplified_model_filepath):
#     """根据设备名称字段匹配终端型号精简6.xlsx表中的终端型号，并填充精简型号字段"""
#     try:
#         # 读取终端型号精简6.xlsx文件
#         #simplified_model_df = pd.read_excel(simplified_model_filepath, sheet_name="Sheet1", engine='openpyxl')
#         simplified_model_df = pd.read_excel(simplified_model_filepath, sheet_name="Sheet2 (2)仅修改HN8145V", engine='openpyxl')
#
#         # 强制转换为字符串类型
#         df['设备名称'] = df['设备名称'].astype(str)
#         simplified_model_df['终端型号'] = simplified_model_df['终端型号'].astype(str)
#
#         # 执行VLOOKUP操作
#         #df = pd.merge(df, simplified_model_df[['终端型号', '精简型号']], left_on='设备名称', right_on='终端型号', how='left')
#         df = pd.merge(df, simplified_model_df[['终端型号', '精简型号2']], left_on='设备名称', right_on='终端型号', how='left')
#
#         # # 将匹配到的精简型号赋值给新列
#         # df['精简型号'] = df['精简型号2']
#         # df.drop(columns=['终端型号', '精简型号2'], inplace=True)
#         df['精简型号'] = df['精简型号2']  # 将 "精简型号2" 的值赋给 "精简型号"
#         df.drop(columns=['终端型号', '精简型号2'], inplace=True)  # 删除不需要的列
#
#
#         return df
#     except KeyError as e:
#         print(f"KeyError: 缺少列：{e}")
#         return None
#     except Exception as e:
#         print(f"VLOOKUP 过程中发生错误：{e}")
#         return None

# #以下这个perform_simplified_model_vlookup 可以实现功能，但是会有点报错
# def perform_simplified_model_vlookup(df, simplified_model_filepath):
#     """根据设备名称字段匹配终端型号精简6.xlsx表中的终端型号，并填充精简型号字段"""
#     try:
#         # 读取终端型号精简6.xlsx文件
#         simplified_model_df = pd.read_excel(simplified_model_filepath, sheet_name="Sheet2 (2)仅修改HN8145V", engine='openpyxl')
#
#         # 彻底清理列名：去除所有非汉字和字母数字字符
#         simplified_model_df.columns = [
#             re.sub(r'[^\w\u4e00-\u9fff]', '', col.strip())
#             for col in simplified_model_df.columns
#         ]
#
#         # 验证列名是否存在
#         required_columns = ['终端型号', '精简型号2']
#         if not all(col in simplified_model_df.columns for col in required_columns):
#             missing = [col for col in required_columns if col not in simplified_model_df.columns]
#             print(f"错误：数据表中缺少关键列 {missing}")
#             return None
#
#         # 强制转换为字符串类型并去除空格
#         df['设备名称'] = df['设备名称'].astype(str).str.strip()
#         simplified_model_df['终端型号'] = simplified_model_df['终端型号'].astype(str).str.strip()
#
#         # 执行 VLOOKUP 操作
#         df = pd.merge(
#             df,
#             simplified_model_df[['终端型号', '精简型号2']],
#             left_on='设备名称',
#             right_on='终端型号',
#             how='left'
#         )
#
#         # 填充新列并删除冗余列
#         df['精简型号'] = df['精简型号2'].fillna('')  # 处理空值
#         df.drop(columns=['终端型号', '精简型号2'], inplace=True, errors='ignore')  # 安全删除列
#
#         return df
#     except KeyError as e:
#         print(f"KeyError: {e}")
#         return None
#     except Exception as e:
#         print(f"VLOOKUP 失败：{e}")
#         return None





# def perform_simplified_model_vlookup(df, simplified_model_filepath):
#     try:
#         # 读取终端型号精简6.xlsx文件
#         simplified_model_df = pd.read_excel(simplified_model_filepath, sheet_name="Sheet2 (2)仅修改HN8145V", engine='openpyxl')
#
#         # 彻底清理列名：去除所有非汉字和字母数字字符
#         simplified_model_df.columns = [
#             re.sub(r'[^\w\u4e00-\u9fff]', '', col.strip())
#             for col in simplified_model_df.columns
#         ]
#
#         # 验证列名是否存在
#         required_columns = ['终端型号', '精简型号2']
#         if not all(col in simplified_model_df.columns for col in required_columns):
#             missing = [col for col in required_columns if col not in simplified_model_df.columns]
#             print(f"错误：数据表中缺少关键列 {missing}")
#             return None
#
#         # 强制转换为字符串类型并去除空格
#         df['设备名称'] = df['设备名称'].astype(str).str.strip()
#         simplified_model_df['终端型号'] = simplified_model_df['终端型号'].astype(str).str.strip()
#
#         # 执行 VLOOKUP 操作
#         df = pd.merge(
#             df,
#             simplified_model_df[['终端型号', '精简型号2']],
#             left_on='设备名称',
#             right_on='终端型号',
#             how='left'
#         )
#
#         # 填充新列并删除冗余列
#         df['精简型号'] = df['精简型号2'].fillna('')  # 处理空值
#         df.drop(columns=['终端型号', '精简型号2'], inplace=True, errors='ignore')  # 安全删除列
#
#         # 保存更新后的数据到目标文件
#         updated_sheet_name = truncate_sheet_name(filtered_sheet_name + "_更新精简型号")
#         with pd.ExcelWriter(filtered_filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#             df.to_excel(writer, sheet_name=updated_sheet_name, index=False)
#
#         print(f"“精简型号”字段已成功更新并保存到 {filtered_filepath} 的 {updated_sheet_name} 工作表中。")
#         return df
#     except KeyError as e:
#         print(f"KeyError: {e}")
#         return None
#     except Exception as e:
#         print(f"VLOOKUP 失败：{e}")
#         return None
# vlookup_module.py 修改后的函数

def perform_simplified_model_vlookup(df, simplified_model_filepath):
    """根据设备名称字段匹配终端型号精简6.xlsx表中的终端型号，并填充精简型号字段"""
    try:
        # 读取终端型号精简6.xlsx文件
        simplified_model_df = pd.read_excel(
            simplified_model_filepath,
            sheet_name="Sheet2 (2)仅修改HN8145V",
            engine='openpyxl'
        )

        # 清理列名
        simplified_model_df.columns = [
            re.sub(r'[^\w\u4e00-\u9fff]', '', col.strip())
            for col in simplified_model_df.columns
        ]

        # 验证必需列是否存在
        required_columns = ['终端型号', '精简型号2']
        if not all(col in simplified_model_df.columns for col in required_columns):
            missing = [col for col in required_columns if col not in simplified_model_df.columns]
            print(f"错误：数据表中缺少关键列 {missing}")
            return None

        # 标准化数据
        df['设备名称'] = df['设备名称'].astype(str).str.strip()
        simplified_model_df['终端型号'] = simplified_model_df['终端型号'].astype(str).str.strip()

        # 执行VLOOKUP
        df = pd.merge(
            df,
            simplified_model_df[['终端型号', '精简型号2']],
            left_on='设备名称',
            right_on='终端型号',
            how='left'
        )

        # 处理合并结果
        df['精简型号'] = df['精简型号2'].fillna('')
        df.drop(columns=['终端型号', '精简型号2'], inplace=True, errors='ignore')

        return df  # 只返回处理后的DataFrame，不处理保存操作

    except Exception as e:
        print(f"处理精简型号时发生错误：{e}")
        return None



# def perform_vlookup_correct(df_target, df_lookup):
#     """
#     使用 pandas.merge 和条件赋值执行正确的 VLOOKUP 操作。
#
#     Args:
#         df_target: 目标 DataFrame（“终端出库报_筛选后1”）。
#         df_lookup: 查找 DataFrame（排序后的 SN 数据）。
#
#     Returns:
#         修改后的目标 DataFrame，或 None 如果发生错误。
#     """
#     try:
#         # 使用 'rms_access_code' 进行左连接
#         merged_df = pd.merge(df_target, df_lookup, left_on="业务号码", right_on="rms_access_code", how="left")
#
#         # 使用条件赋值，仅在匹配成功时更新 "LOID（SN码）" 列
#         df_target["LOID（SN码）"] = merged_df["ce_loid"]
#
#         return df_target
#
#     except KeyError as e:
#         print(f"KeyError: 列名 '{e.args[0]}' 不存在。请检查 DataFrame 的列名。")
#         return None
#     except Exception as e:
#         print(f"VLOOKUP 操作失败：{e}")
#         return None

#
# def perform_vlookup(df, lookup_df, lookup_col='rms_access_code', result_col='ce_loid', new_col_name='LOID（SN码）'):
#     """
#     在 DataFrame 中执行单列 VLOOKUP 操作。
#
#     参数：
#         df (pd.DataFrame): 主 DataFrame，要在其中添加新列。
#         lookup_df (pd.DataFrame): 查找 DataFrame，包含查找值和结果值。
#         lookup_col (str): 查找 DataFrame 中用于查找的列名，默认为 'rms_access_code'。
#         result_col (str): 查找 DataFrame 中要返回的结果列名，默认为 'ce_loid'。
#         new_col_name (str): 在主 DataFrame 中创建的新列的名称，默认为 'LOID（SN码）'。
#
#     返回值：
#         pd.DataFrame: 修改后的主 DataFrame，如果发生错误则返回 None。
#     """
#     try:
#         if lookup_col not in lookup_df.columns:
#             print(f"错误：查找 DataFrame 中不存在列 '{lookup_col}'。")
#             return None
#         if result_col not in lookup_df.columns:
#             print(f"错误：查找 DataFrame 中不存在列 '{result_col}'。")
#             return None
#
#         # 将查找 DataFrame 的查找列设置为索引，以提高查找效率
#         lookup_df = lookup_df.set_index(lookup_col)
#
#         # 使用 map 函数执行查找
#         df[new_col_name] = df[lookup_col].map(lookup_df[result_col])
#         print(f"成功在 DataFrame 中执行单列 VLOOKUP 操作，新列名为 '{new_col_name}'。")
#         return df
#
#     except KeyError as e:
#         print(f"错误：主 DataFrame 中不存在列 '{lookup_col}'。错误信息：{e}")
#         return None
#     except Exception as e:
#         print(f"执行单列 VLOOKUP 操作时发生未知错误：{e}")
#         return None
#
# def perform_vlookup_multi(df, lookup_df, lookup_cols, result_col='ce_loid', new_col_name='LOID（SN码）'):
#     """
#     根据多个查找列在 DataFrame 中执行 VLOOKUP 操作。
#
#     参数：
#         df (pd.DataFrame): 主 DataFrame，要在其中添加新列。
#         lookup_df (pd.DataFrame): 查找 DataFrame，包含查找值和结果值。
#         lookup_cols (list): 查找 DataFrame 中用于查找的列名列表。
#         result_col (str): 查找 DataFrame 中要返回的结果列名，默认为 'ce_loid'。
#         new_col_name (str): 在主 DataFrame 中创建的新列的名称，默认为 'LOID（SN码）'。
#
#     返回值：
#         pd.DataFrame: 修改后的主 DataFrame，如果发生错误则返回 None。
#     """
#     try:
#         # 检查查找列是否存在
#         for col in lookup_cols:
#             if col not in df.columns:
#                 print(f"错误：主 DataFrame 中不存在列 '{col}'。")
#                 return None
#             if col not in lookup_df.columns:
#                 print(f"错误：查找 DataFrame 中不存在列 '{col}'。")
#                 return None
#         if result_col not in lookup_df.columns:
#             print(f"错误：查找 DataFrame 中不存在列 '{result_col}'。")
#             return None
#
#         # 创建一个用于合并的键，使用astype(str)处理不同数据类型
#         lookup_df['merge_key'] = lookup_df[lookup_cols].apply(lambda x: '_'.join(x.astype(str)), axis=1)
#         df['merge_key'] = df[lookup_cols].apply(lambda x: '_'.join(x.astype(str)), axis=1)
#
#         # 避免 SettingWithCopyWarning
#         lookup_df = lookup_df.copy()
#         df = df.copy()
#
#         lookup_df = lookup_df.set_index('merge_key')
#
#         df[new_col_name] = df['merge_key'].map(lookup_df[result_col])
#
#         df = df.drop(columns=['merge_key'])
#         lookup_df = lookup_df.reset_index(drop=True)
#
#         print(f"成功在 DataFrame 中执行多列 VLOOKUP 操作，新列名为 '{new_col_name}'。")
#         return df
#
#     except KeyError as e:
#         print(f"错误：主 DataFrame 或查找 DataFrame 中缺少列。错误信息：{e}")
#         return None
#     except Exception as e:
#         print(f"执行多列 VLOOKUP 操作时发生未知错误：{e}")
#         return None
#
