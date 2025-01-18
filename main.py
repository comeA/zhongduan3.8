import os
import pandas as pd
from modules.sn_data_process import process_sn_data, sort_and_save_sn_data
from modules.vlookup_module import perform_vlookup_correct
from modules.insert_columns import insert_columns
from modules.excel_utils import copy_data_to_excel # 导入新的模块
from modules.business_number_utils import copy_business_numbers_to_template # 导入新的模块




def get_sheet_name(prompt):
    """循环提示用户输入工作表名称，直到输入非空值为止。"""
    while True:
        sheet_name = input(prompt).strip()
        if sheet_name:
            return sheet_name
        else:
            print("工作表名称不能为空，请重新输入。")

def get_file_path(prompt, check_exists=True):
    """循环提示用户输入文件路径，直到满足条件为止。"""
    while True:
        file_path = input(prompt).replace("\\", "/").strip()
        if not file_path:
            print("文件路径不能为空，请重新输入。")
            continue
        if check_exists and not os.path.exists(file_path):
            print(f"文件路径 {file_path} 不存在，请重新输入。")
        else:
            return file_path

def get_file_name(prompt):
    """循环提示用户输入文件名，直到输入非空值为止。"""
    while True:
        file_name = input(prompt).strip()
        if file_name:
            return file_name
        else:
            print("文件名不能为空，请重新输入。")

def get_yn_input(prompt):
    """循环提示用户输入y/n，直到输入正确为止"""
    while True:
        user_input = input(prompt).strip().lower()
        if user_input in ('y', 'n'):
            return user_input
        else:
            print("请输入 y 或 n。")



if __name__ == "__main__":
   # print(f"当前工作目录：{os.getcwd()}")  # 测试数据导入的文件是否存在importCPEID2ExportSN.xlsx  和 importCPEID2ExportMAC.xlsx

    print("欢迎使用终端数据处理程序！")

    # 源文件输入
    source_folder = get_file_path("请输入源文件所在文件夹路径：")
    source_filename = get_file_name("请输入源文件名（包含扩展名，例如：表05终端工单一览表.xlsx）：")
    source_filepath = os.path.join(source_folder, source_filename)
    if not os.path.exists(source_filepath):
        print(f"源文件 {source_filepath} 不存在，程序退出。")
        exit()

    source_sheet = get_sheet_name("请输入源文件子表名称：")

    # 目标文件输入
    dest_folder = get_file_path("请输入目标文件所在文件夹路径（可新建）：", check_exists=False)
    os.makedirs(dest_folder, exist_ok=True)
    dest_filename = get_file_name("请输入目标文件名（包含扩展名，例如：终端出库报.xlsx）：")
    dest_filepath = os.path.join(dest_folder, dest_filename)

    result_sheet = get_sheet_name("请输入目标文件子表名称：")

    process_result, filtered_sheet_name, filtered_filepath = process_sn_data(source_filepath, source_sheet, dest_filepath, result_sheet)

    if process_result:
        try:
            df = pd.read_excel(filtered_filepath, sheet_name=filtered_sheet_name, engine='openpyxl')
        except Exception as e:
            print(f"读取筛选后的excel失败：{e}")
            exit()

            # *** 正确处理并复制“业务号码”的代码 ***
            while True:
                import_template_filepath = get_file_path("请输入“导入模板.xlsx”文件路径：")
                if os.path.splitext(import_template_filepath)[1].lower() != ".xlsx":
                    print("文件格式不正确，请选择xlsx文件")
                    continue
                if not os.path.exists(import_template_filepath):
                    print("文件不存在，请检查路径")
                    continue
                break
            if not copy_business_numbers_to_template(df, import_template_filepath): #调用新的函数
                exit()
            # *** 修改部分结束 ***

        insert_cols_before_vlookup = get_yn_input("数据已成功复制和筛选。是否立即在筛选后的sheet中插入新列？(y/n): ")
        if insert_cols_before_vlookup == 'y':
            insert_after_sheet_name = filtered_sheet_name + "_插入后1"
            insert_result, df = insert_columns(filtered_filepath, filtered_sheet_name, insert_after_sheet_name,
                                               df)  # 接收df
            if insert_result:
                print("筛选后的sheet新列插入成功！")
                df.to_excel(filtered_filepath, sheet_name=insert_after_sheet_name, index=False,
                            engine='openpyxl')  # 保存到新的sheet
                try:
                    print(f"插入新列后的DataFrame列名：{df.columns.tolist()}")
                    filtered_sheet_name = insert_after_sheet_name
                except Exception as e:
                    print(f"重新读取excel失败：{e}")
                    exit()
            else:
                print("筛选后的sheet新列插入失败！")
                exit()
        else:
            print("跳过插入新列操作。")

        continue_processing = get_yn_input("是否继续处理“业务号码-LOID（SN码）”数据？(y/n): ")
        if continue_processing == 'y':
            while True:
                sn_data_filepath = get_file_path("请输入“业务号码-LOID（SN码）”数据文件路径（例如：终端数据匹配逻辑1sn码.xlsx 或 终端数据匹配逻辑1sn码.txt）：")
                sn_sheet = get_sheet_name("请输入“业务号码-LOID（SN码）”文件子表名称：")

                # 判断文件类型
                if sn_data_filepath.lower().endswith(".txt"):
                    special_format = True
                elif sn_data_filepath.lower().endswith((".xls", ".xlsx")):
                    special_format = False
                else:
                    print("不支持的文件类型，请选择txt或excel文件")
                    continue

                if "dwd_hzluheb_acc_sn_final_pg" not in os.path.basename(sn_data_filepath).lower() and "dwd_hzluheb_acc_sn_final_pg" not in sn_sheet.lower():
                    print("文件名或子表名不包含关键字 dwd_hzluheb_acc_sn_final_pg，请检查")
                    continue

                sorted_sn_filepath = os.path.join(os.path.dirname(sn_data_filepath), "sorted_" + os.path.basename(sn_data_filepath))

                sort_result = sort_and_save_sn_data(sn_data_filepath, sn_sheet, sorted_sn_filepath, special_format)
                if not sort_result:
                    print("处理SN数据失败，请检查文件内容和格式。")
                    continue

                try:
                    sn_df = pd.read_excel(sorted_sn_filepath, sheet_name=sn_sheet, engine='openpyxl')

                    # *** 正确的 VLOOKUP 逻辑，直接使用更新后的 df ***
                    if "业务号码" in df.columns and "rms_access_code" in sn_df.columns and "ce_loid" in sn_df.columns and "LOID（SN码）" in df.columns:
                        df = perform_vlookup_correct(df, sn_df)
                        if df is not None:
                            try:
                                match_sn_sheet_name = insert_after_sheet_name + "_匹配SN" # 新的 sheet 名称
                                with pd.ExcelWriter(filtered_filepath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                                    df.to_excel(writer, sheet_name=match_sn_sheet_name, index=False) # 使用新的 sheet 名称
                                print(f"LOID（SN码）已成功匹配并添加到文件 {filtered_filepath} 的 {match_sn_sheet_name} 工作表中。")
                                filtered_sheet_name = match_sn_sheet_name #更新filtered_sheet_name，方便后续操作

                                # *** 修改部分开始 ***
                                # 复制数据到 importCPEID2Export*.xlsx 文件
                                while True:
                                    sn_export_filepath = get_file_path("请输入importCPEID2ExportSN.xlsx文件路径：")
                                    if os.path.splitext(sn_export_filepath)[1].lower() != ".xlsx":
                                        print("文件格式不正确，请选择xlsx文件")
                                        continue
                                    if not os.path.exists(sn_export_filepath):
                                        print("文件不存在，请检查路径")
                                        continue
                                    break

                                while True:
                                    mac_export_filepath = get_file_path("请输入importCPEID2ExportMAC.xlsx文件路径：")
                                    if os.path.splitext(mac_export_filepath)[1].lower() != ".xlsx":
                                        print("文件格式不正确，请选择xlsx文件")
                                        continue
                                    if not os.path.exists(mac_export_filepath):
                                        print("文件不存在，请检查路径")
                                        continue
                                    break

                                if not copy_data_to_excel(df['LOID（SN码）'], sn_export_filepath, "Sheet1", "LOID"):
                                    continue
                                if not copy_data_to_excel(df['ISCM终端MAC地址'], mac_export_filepath, "Sheet1", "CPEID(OUI和序列号必须填写)"):
                                    continue
                                # *** 修改部分结束 ***

                            except Exception as e:
                                print(f"保存匹配结果到目标文件时发生错误：{e}")
                        else:
                            print("VLOOKUP 操作失败。")
                    else:
                        print(f"错误：目标 DataFrame 中缺少列 '业务号码' 或 SN DataFrame 中缺少列 'rms_access_code' 或 'ce_loid' 或目标DataFrame中缺少‘LOID（SN码）’。")
                    break #vlookup成功后退出循环
                except FileNotFoundError as e:
                    print(f"文件不存在：{e}，请检查文件是否存在")
                except ValueError as e:
                    print(f"工作表不存在或文件格式错误：{e}，请检查工作表是否存在")
                except Exception as e:
                    print(f"其他错误：{e}")

        elif continue_processing.lower() == 'n':
            print("操作完成！")
        else:
            print("无效的输入，操作完成！")
    else:
        print("处理失败！")

    print("程序结束。")