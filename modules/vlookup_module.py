import pandas as pd

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
