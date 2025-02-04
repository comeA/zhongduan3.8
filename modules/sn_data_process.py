# import pandas as pd
# import os
# import re
# from modules.copy_sheet import copy_sheet_data
import pandas as pd
import os
import re
import chardet # 导入chardet
import openpyxl
from modules.copy_sheet import copy_sheet_data

def process_special_format_data(data_string):
    """处理特殊格式的数据字符串，允许create_date为空"""
    try:
        match = re.match(r"(.*),(.*),'(.*)'", data_string) or re.match(r"(.*),(.*),", data_string)
        if match:
            rms_access_code = match.group(1).strip()
            ce_loid = match.group(2).strip()
            create_date_str = match.group(3).strip() if match.lastindex == 3 else None

            if create_date_str:
                try:
                    create_date = pd.to_datetime(create_date_str, format='%Y-%m-%d')
                except ValueError as e:
                    print(f"日期转换错误: {e}，原始字符串为：{create_date_str}")
                    create_date = None
            else:
                create_date = None

            return {'rms_access_code': rms_access_code, 'ce_loid': ce_loid, 'create_date': create_date}
        else:
            print(f"数据格式不匹配: {data_string}")
            return None
    except Exception as e:
        print(f"处理数据时发生错误：{e}")
        return None

def sort_and_save_sn_data(filepath, sheet_name, save_filepath, special_format=False):
    """对SN数据进行排序并保存。"""
    print("-" * 30)
    print("开始执行 sort_and_save_sn_data 函数")
    print(f"原始文件路径：{filepath}")
    print(f"原始工作表：{sheet_name}")
    print(f"保存文件路径：{save_filepath}")
    print(f"是否为特殊格式：{special_format}")
    try:
        if special_format:  # 处理特殊格式的文本文件
            data = []
            try:
                with open(filepath, 'r', encoding='utf-8') as f:  # 指定编码方式
                    for line in f:
                        processed_data = process_special_format_data(line.strip())
                        if processed_data:
                            data.append(processed_data)
            except FileNotFoundError:
                print(f"错误：文件 {filepath} 未找到。")
                return False
            except Exception as e:
                print(f"读取或处理文件时出错：{e}")
                return False
            df = pd.DataFrame(data)
        else:  # 处理 Excel 文件
            try:
                df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
            except FileNotFoundError:
                print(f"错误：文件 {filepath} 未找到。")
                return False
            except KeyError:
                print(f"错误：工作表 {sheet_name} 未找到。")
                return False
            except Exception as e:
                print(f"读取文件失败：{e}")
                return False

        if "create_date" not in df.columns:
            print("错误：SN数据文件中缺少 'create_date' 列，无法进行排序。")
            return False
        # 降序排序，这是匹配的基础！
        df_sorted = df.sort_values(by="create_date", ascending=False)


        try:
            df_sorted.to_excel(save_filepath, index=False, sheet_name=sheet_name, engine='openpyxl')
            print(f"SN数据已按 create_date 降序排序并保存到：{save_filepath} 的 {sheet_name} 工作表")
            return True
        except Exception as e:
            print(f"保存排序后的文件时发生错误：{e}")
            return False

    except Exception as e:
        print(f"sort_and_save_sn_data 函数执行过程中发生未知错误：{e}")
        return False
    finally:
        print("sort_and_save_sn_data 函数执行完毕")
        print("-" * 30)



def process_sn_data(source_filepath, source_sheet, dest_filepath, result_sheet):
    """从源文件复制数据到目标文件，并进行筛选。"""
    print("-" * 30)
    print("开始执行 process_sn_data 函数")
    print(f"源文件路径：{source_filepath}, 源工作表：{source_sheet}")
    print(f"目标文件路径：{dest_filepath}, 结果工作表：{result_sheet}")

    try:
        copy_result, new_sheet_name = copy_sheet_data(source_filepath, source_sheet, dest_filepath, result_sheet)
        if not copy_result:
            print("复制数据失败，无法进行后续操作。")
            return False, None, None

        try:
            df = pd.read_excel(dest_filepath, sheet_name=new_sheet_name, engine='openpyxl')
        except FileNotFoundError:
            print(f"错误：文件 {dest_filepath} 未找到。")
            return False, None, None
        except KeyError:
            print(f"错误：工作表 {new_sheet_name} 未找到。")
            return False, None, None
        except Exception as e:
            print(f"读取excel失败：{e}")
            return False, None, None

        # *** 正确的筛选逻辑 ***
        print("开始筛选数据...")
        mask = (df['终端4级目录名称'].str.contains('A8C', na=False) | df['终端4级目录名称'].str.contains('光猫', na=False)) & \
               (df['业务工单回单类型'].str.contains('在途', na=False)) & \
               (~df['条形码'].isin(['', "'", "''", " ", None]))
        df_filtered = df[mask].copy()  # 加上.copy()，避免SettingWithCopyWarning
        print("数据筛选完成。")

        # *** 保存筛选后的数据到新的工作表 ***
        #filtered_sheet_name = new_sheet_name + "_筛选后"
        filtered_sheet_name = new_sheet_name + "_筛选后1"  # 修改这里
        try:
            with pd.ExcelWriter(dest_filepath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df_filtered.to_excel(writer, sheet_name=filtered_sheet_name, index=False)
        except Exception as e:
            print(f"保存筛选后的数据时发生错误：{e}")
            return False, None, None

        print(f"筛选后的数据已保存到 {dest_filepath} 的 {filtered_sheet_name} 工作表。")
        return True, filtered_sheet_name, dest_filepath

    except Exception as e:
        print(f"process_sn_data 函数执行过程中发生未知错误：{e}")
        return False, None, None
    finally:
        print("process_sn_data 函数执行完毕")
        print("-" * 30)

'''
import pandas as pd
import os
from modules.copy_sheet import copy_sheet_data

def sort_and_save_sn_data(filepath, sheet_name, save_filepath):
    """读取SN数据，根据rms_access_code排序并保存到新文件"""

    print("-" * 30)
    print(f"开始执行 sort_and_save_sn_data 函数")
    print(f"原始文件路径：{filepath}")
    print(f"保存文件路径：{save_filepath}")

    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
        if df.empty:
            print(f"警告：文件{filepath}的{sheet_name}工作表为空，无法进行后续操作。")
            return False

        if 'rms_access_code' not in df.columns:
            print("错误：SN数据文件中缺少 rms_access_code 列，无法排序。请检查文件格式。")
            return False

        df_sorted = df.sort_values(by='rms_access_code')

        save_dir = os.path.dirname(save_filepath)
        os.makedirs(save_dir, exist_ok=True)  # 创建目录，如果存在则不报错

        df_sorted.to_excel(save_filepath, index=False, engine='openpyxl')
        print(f"SN数据已排序并保存到：{save_filepath}")
        return True

    except FileNotFoundError:
        print(f"错误：文件 {filepath} 未找到。")
        return False
    except ValueError as e:
        print(f"错误：读取文件时发生值错误：{e}，请检查工作表名称是否正确。")
        return False
    except Exception as e:
        print(f"排序并保存文件时发生未知错误：{e}")
        return False
    finally:
        print(f"sort_and_save_sn_data 函数执行完毕")
        print("-" * 30)


def process_sn_data(source_filepath, source_sheet, dest_filepath, result_sheet):
    """处理主流程数据：先复制，后在目标文件上按正确条件筛选，并保存到新子表"""
    print("-" * 30)
    print("开始执行 process_sn_data 函数")
    print(f"源文件路径：{source_filepath}, 源工作表：{source_sheet}")
    print(f"目标文件路径：{dest_filepath}, 结果工作表：{result_sheet}")

    try:
        # 检查目标文件是否存在，如果不存在则创建
        if not os.path.exists(dest_filepath):
            try:
                open(dest_filepath, 'w').close()
                print(f"目标文件 {dest_filepath} 不存在，已创建空文件。")
            except Exception as e:
                print(f"创建目标文件失败：{e}")
                return False, None

        if not copy_sheet_data(source_filepath, source_sheet, dest_filepath, result_sheet):
            return False, None

        try:
            df = pd.read_excel(dest_filepath, sheet_name=result_sheet, engine='openpyxl')
        except ValueError as e:
            print(f"读取目标文件数据失败：{e}。请检查工作表 '{result_sheet}' 是否存在。")
            return False, None
        except FileNotFoundError as e:
            print(f"读取目标文件数据失败：{e}。请检查文件是否存在")
            return False, None
        except Exception as e:
            print(f"读取目标文件时发生其他错误：{e}")
            return False, None

        if df.empty:
            print(f"警告：目标文件{dest_filepath}的{result_sheet}工作表为空，无法进行后续操作。")
            return False, None

        mask = (df['终端4级目录名称'].str.contains('A8C', na=False) | df['终端4级目录名称'].str.contains('光猫', na=False)) & \
               (df['业务工单回单类型'].str.contains('在途', na=False)) & \
               (~df['条形码'].isin(['', "'", "''", " ", None]))

        filtered_df = df[mask].copy()

        new_sheet_name = result_sheet + "_筛选后"
        try:
            with pd.ExcelWriter(dest_filepath, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:  # 修改为 'new'
                filtered_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
            print(f"筛选后的数据已成功保存到工作表：{new_sheet_name}")
        except Exception as e:
            print(f"保存筛选后的数据时发生错误：{e}")
            return False, None

        return True, new_sheet_name  # 返回 True 和新 sheet 名称

    except Exception as e:
        print(f"process_sn_data 函数执行过程中发生未知错误：{e}")
        return False, None
    finally:
        print("process_sn_data 函数执行完毕")
        print("-" * 30)
'''