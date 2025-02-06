# Excel_处理终端数据 V3.8/main.py

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


import os
import pandas as pd
import openpyxl
import sys

# 导入各个模块，并使用别名
import modules.excel_utils as excel_utils
import modules.business_number_utils as business_number_utils
import modules.sn_data_process as sn_data_process
import modules.vlookup_module as vlookup_module
import modules.insert_columns as insert_columns
import modules.utils as utils
import modules.copy_to_template as copy_to_template
import modules.cpe_processing as cpe_processing  # 导入 cpe_processing 模块

# 从 utils 模块导入需要的函数，并赋予更具描述性的别名
get_file_path = utils.get_file_path
get_file_name = utils.get_file_name
get_sheet_name = utils.get_sheet_name
get_yes_no_input = utils.get_yes_no_input
create_directory = utils.create_directory
check_file_exists = utils.check_file_exists
check_excel_file = utils.check_excel_file

# 重新赋值，保持命名一致性
copy_data_to_excel = excel_utils.copy_data_to_excel
copy_data_to_excel_with_header = excel_utils.copy_data_to_excel_with_header
copy_data_to_excel_with_header_split_columns = excel_utils.copy_data_to_excel_with_header_split_columns

copy_business_numbers_to_template = business_number_utils.copy_business_numbers_to_template
process_sn_data = sn_data_process.process_sn_data
sort_and_save_sn_data = sn_data_process.sort_and_save_sn_data
perform_vlookup_correct = vlookup_module.perform_vlookup_correct
insert_columns_func = insert_columns.insert_columns  # 修改变量名，避免与模块名冲突

print(f"Python 解释器路径：{sys.executable}")
print(f"openpyxl 版本：{openpyxl.__version__}")

def main():
    print("欢迎使用终端数据处理程序！")

    # 源文件输入
    source_folder = get_file_path("请输入源文件所在文件夹路径：")
    source_filename = get_file_name("请输入源文件名（包含扩展名）：")
    source_filepath = os.path.join(source_folder, source_filename)

    if not check_file_exists(source_filepath):
        print(f"源文件 {source_filepath} 不存在，程序退出。")
        return

    source_sheet = get_sheet_name("请输入源文件子表名称：")

    # 目标文件输入
    dest_folder = get_file_path("请输入目标文件所在文件夹路径（可新建）：", check_exists=False)
    create_directory(dest_folder)

    dest_filename = get_file_name("请输入目标文件名（包含扩展名）：")
    dest_filepath = os.path.join(dest_folder, dest_filename)

    result_sheet = get_sheet_name("请输入目标文件子表名称：")

    process_result, filtered_sheet_name, filtered_filepath = process_sn_data(source_filepath, source_sheet, dest_filepath, result_sheet)

    if process_result:
        try:
            df = pd.read_excel(filtered_filepath, sheet_name=filtered_sheet_name, engine='openpyxl')
        except Exception as e:
            print(f"读取筛选后的 Excel 文件失败：{e}")
            return

        # 复制业务号码到导入模板
        import_template_filepath = get_file_path("请输入“导入模板.xlsx”文件路径：")
        if not copy_business_numbers_to_template(df, import_template_filepath):
            print("复制业务号码失败")
            return

        insert_cols_before_vlookup = get_yes_no_input("数据已成功复制。是否立即在筛选后的 sheet 中插入新列？(y/n): ")

        if insert_cols_before_vlookup == 'y':
            insert_after_sheet_name = filtered_sheet_name + "_插入后1"
            insert_result, df = insert_columns_func(filtered_filepath, filtered_sheet_name, insert_after_sheet_name, df)  # 调用修改后的函数名

            if insert_result:
                print("筛选后的 sheet 新列插入成功！")
                df.to_excel(filtered_filepath, sheet_name=insert_after_sheet_name, index=False, engine='openpyxl')
                filtered_sheet_name = insert_after_sheet_name
            else:
                print("筛选后的 sheet 新列插入失败！")
                return
        else:
            print("跳过插入新列操作。")

        continue_processing = get_yes_no_input("是否继续处理“业务号码-LOID（SN 码）”数据？(y/n): ")

        if continue_processing == 'y':
            sn_data_filepath = get_file_path("请输入“业务号码-LOID（SN 码）”数据文件路径：")
            sn_sheet = get_sheet_name("请输入“业务号码-LOID（SN 码）”文件子表名称：")

            if not (sn_data_filepath.lower().endswith((".txt", ".xls", ".xlsx"))):
                print("不支持的文件类型，请选择 txt 或 Excel 文件")
                return

            if "dwd_hzluheb_acc_sn_final_pg" not in os.path.basename(
                    sn_data_filepath).lower() and "dwd_hzluheb_acc_sn_final_pg" not in sn_sheet.lower():
                print("文件名或子表名不包含关键字 dwd_hzluheb_acc_sn_final_pg，请检查")
                return

            sorted_sn_filepath = os.path.join(os.path.dirname(sn_data_filepath),
                                              "sorted_" + os.path.basename(sn_data_filepath))

            if not sort_and_save_sn_data(sn_data_filepath, sn_sheet, sorted_sn_filepath,
                                         sn_data_filepath.lower().endswith(".txt")):
                print("处理 SN 数据失败，请检查文件内容和格式。")
                return

            try:
                sn_df = pd.read_excel(sorted_sn_filepath, sheet_name=sn_sheet, engine='openpyxl')

                if all(col in df.columns for col in ["业务号码", "LOID（SN码）"]) and all(
                        col in sn_df.columns for col in ["rms_access_code", "ce_loid"]):
                    df = perform_vlookup_correct(df, sn_df)
                    if df is not None:
                        try:
                            match_sn_sheet_name = filtered_sheet_name + "_匹配SN"
                            with pd.ExcelWriter(filtered_filepath, engine='openpyxl', mode='a',
                                                if_sheet_exists='overlay') as writer:
                                df.to_excel(writer, sheet_name=match_sn_sheet_name, index=False)
                            print(f"LOID（SN 码）已成功匹配并添加到文件 {filtered_filepath} 的 {match_sn_sheet_name} 工作表中。")
                            filtered_sheet_name = match_sn_sheet_name

                            # 复制数据到 importCPEID2Export*.xlsx 文件
                            while True:  # 循环获取正确的文件路径
                                sn_export_filepath = get_file_path("请输入 importCPEID2ExportSN.xlsx 文件路径：")
                                if os.path.splitext(sn_export_filepath)[1].lower() != ".xlsx":
                                    print("文件格式不正确，请选择xlsx文件")
                                    continue
                                if not os.path.exists(sn_export_filepath):
                                    print("文件不存在，请检查路径")
                                    continue
                                break
                            while True:  # 循环获取正确的文件路径
                                mac_export_filepath = get_file_path("请输入 importCPEID2ExportMAC.xlsx 文件路径：")
                                if os.path.splitext(mac_export_filepath)[1].lower() != ".xlsx":
                                    print("文件格式不正确，请选择xlsx文件")
                                    continue
                                if not os.path.exists(mac_export_filepath):
                                    print("文件不存在，请检查路径")
                                    continue
                                break

                            if not copy_data_to_excel_with_header(df['LOID（SN码）'], sn_export_filepath, "Sheet1",
                                                                  "LOID"):
                                print("复制 LOID（SN 码）到 importCPEID2ExportSN.xlsx 失败")
                                return

                            if not copy_data_to_excel_with_header(df['ISCM终端MAC地址'], mac_export_filepath, "Sheet1",
                                                                  "CPEID(OUI和序列号必须填写)"):
                                print("复制 ISCM 终端 MAC 地址到 importCPEID2ExportMAC.xlsx 失败")
                                return

                            ##处理 cpeExport 数据并更新终端出库报表 (移到这里)
                            continue_cpe_processing = get_yes_no_input("是否继续处理目前在用型号2？(y/n): ")
                            if continue_cpe_processing == 'y':
                                cpe_export_filepath = get_file_path("请输入 cpeExport 文件路径：")
                                if cpe_processing.process_cpe_export_data(cpe_export_filepath):
                                    print("cpeExport 数据处理成功！")
                                else:
                                    print("cpeExport 数据处理失败！")

                                if cpe_processing.update_terminal_data(filtered_filepath, cpe_export_filepath):
                                    print("终端出库报表数据更新成功！")
                                else:
                                    print("终端出库报表数据更新失败！")
                            else:
                                print("跳过处理目前在用型号2的操作。")

                            ##处理 cpeExport 数据并更新终端出库报表 (移到这里)

                        except Exception as e:
                            print(f"保存匹配结果到目标文件时发生错误：{e}")
                            import traceback  # 导入 traceback 模块
                            traceback.print_exc()  # 打印详细错误信息
                            return  # 出现错误时，使用 return 退出函数
                    else:
                        print("VLOOKUP 操作失败。请检查源数据和 SN 码数据是否包含所有必需的列。")  # 提示用户检查数据
                        return  # 出现错误时，使用 return 退出函数
                else:
                    print("目标 DataFrame 或 SN DataFrame 缺少必需的列，无法执行 VLOOKUP。")  # 提示用户缺少必需的列
                    return  # 出现错误时，使用 return 退出函数

            except FileNotFoundError as e:
                print(f"文件不存在：{e}，请检查文件是否存在")
                return  # 出现错误时，使用 return 退出函数
            except ValueError as e:
                print(f"工作表不存在或文件格式错误：{e}，请检查工作表是否存在")
                return  # 出现错误时，使用 return 退出函数
            except Exception as e:
                print(f"其他错误：{e}")
                import traceback  # 导入 traceback 模块
                traceback.print_exc()  # 打印详细错误信息
                return  # 出现错误时，使用 return 退出函数


if __name__ == "__main__":
    main()






#
# '''
# import os
# import pandas as pd
# #from modules import excel_utils, business_number_utils, sn_data_process, vlookup_module, insert_columns, utils, \
# from modules import  business_number_utils, sn_data_process, vlookup_module, insert_columns, utils, \
#     copy_to_template
# import sys
# import openpyxl
# import modules.excel_utils as excel_utils # 显式导入，并使用别名
#
# from modules.excel_utils import copy_data_to_excel, copy_data_to_excel_with_header
#
# # 使用更简洁的导入方式，直接将函数名导入到当前命名空间
# get_sheet_name = utils.get_sheet_name
# get_file_path = utils.get_file_path
# get_file_name = utils.get_file_name
# get_yn_input = utils.get_yn_input
# create_directory = utils.create_directory
# check_file_exists = utils.check_file_exists
# check_excel_file = utils.check_excel_file
#
# copy_data_to_excel = excel_utils.copy_data_to_excel
# copy_business_numbers_to_template = business_number_utils.copy_business_numbers_to_template
# process_sn_data = sn_data_process.process_sn_data
# sort_and_save_sn_data = sn_data_process.sort_and_save_sn_data
# perform_vlookup_correct = vlookup_module.perform_vlookup_correct
# insert_columns = insert_columns.insert_columns
#
# '''
# import os
# import pandas as pd
# #from modules import excel_utils, business_number_utils, sn_data_process, vlookup_module, insert_columns, utils, \
# from modules import  business_number_utils, sn_data_process, vlookup_module, insert_columns, utils, \
#     copy_to_template
# import sys
# import openpyxl
# import modules.excel_utils as excel_utils # 显式导入，并使用别名
# from modules.excel_utils import copy_data_to_excel, copy_data_to_excel_with_header, copy_business_numbers_to_excel_with_header_split_columns # 导入新的函数
# from modules.excel_utils import copy_data_to_excel, copy_data_to_excel_with_header, copy_data_to_excel_with_header_split_columns
#
#
# # Use absolute imports for individual functions from utils
# from modules.utils import get_file_path, get_file_name, get_yn_input, create_directory, check_file_exists, check_excel_file
#
# copy_data_to_excel = excel_utils.copy_data_to_excel
# copy_business_numbers_to_template = business_number_utils.copy_business_numbers_to_template
# process_sn_data = sn_data_process.process_sn_data
# sort_and_save_sn_data = sn_data_process.sort_and_save_sn_data
# perform_vlookup_correct = vlookup_module.perform_vlookup_correct
# insert_columns = insert_columns.insert_columns
#
#
# print(f"Python 解释器路径：{sys.executable}")
# print(f"openpyxl 版本：{openpyxl.__version__}")
#
#
# def get_sheet_name(prompt):
#     """循环提示用户输入工作表名称，直到输入非空值为止。"""
#     while True:
#         sheet_name = input(prompt).strip()
#         if sheet_name:
#             return sheet_name
#         else:
#             print("工作表名称不能为空，请重新输入。")
#
# def get_file_path(prompt, check_exists=True):
#     """循环提示用户输入文件路径，直到满足条件为止。"""
#     while True:
#         file_path = input(prompt).replace("\\", "/").strip()
#         if not file_path:
#             print("文件路径不能为空，请重新输入。")
#             continue
#         if check_exists and not os.path.exists(file_path):
#             print(f"文件路径 {file_path} 不存在，请重新输入。")
#         else:
#             return file_path
#
# def get_file_name(prompt):
#     """循环提示用户输入文件名，直到输入非空值为止。"""
#     while True:
#         file_name = input(prompt).strip()
#         if file_name:
#             return file_name
#         else:
#             print("文件名不能为空，请重新输入。")
#
# def get_yn_input(prompt):
#     """循环提示用户输入y/n，直到输入正确为止"""
#     while True:
#         user_input = input(prompt).strip().lower()
#         if user_input in ('y', 'n'):
#             return user_input
#         else:
#             print("请输入 y 或 n。")
#
#
#
# if __name__ == "__main__":
#    # print(f"当前工作目录：{os.getcwd()}")  # 测试数据导入的文件是否存在importCPEID2ExportSN.xlsx  和 importCPEID2ExportMAC.xlsx
#
#     print("欢迎使用终端数据处理程序！")
#
#
#
#
#     # 源文件输入
#     source_folder = get_file_path("请输入源文件所在文件夹路径：")
#     source_filename = get_file_name("请输入源文件名（包含扩展名，例如：表05终端工单一览表.xlsx）：")
#     source_filepath = os.path.join(source_folder, source_filename)
#     if not os.path.exists(source_filepath):
#         print(f"源文件 {source_filepath} 不存在，程序退出。")
#         exit()
#
#     source_sheet = get_sheet_name("请输入源文件子表名称：")
#
#     # 目标文件输入
#     dest_folder = get_file_path("请输入目标文件所在文件夹路径（可新建）：", check_exists=False)
#     os.makedirs(dest_folder, exist_ok=True)
#     dest_filename = get_file_name("请输入目标文件名（包含扩展名，例如：终端出库报.xlsx）：")
#     dest_filepath = os.path.join(dest_folder, dest_filename)
#
#     result_sheet = get_sheet_name("请输入目标文件子表名称：")
#
#     process_result, filtered_sheet_name, filtered_filepath = process_sn_data(source_filepath, source_sheet, dest_filepath, result_sheet)
#
#     if process_result:
#         try:
#             df = pd.read_excel(filtered_filepath, sheet_name=filtered_sheet_name, engine='openpyxl')
#         except Exception as e:
#             print(f"读取筛选后的excel失败：{e}")
#             exit()
#
#
#             # *** 正确处理并复制“业务号码”的代码 ***
#
#         #     # 复制业务号码到导入模板
#         # while True:
#         #     import_template_filepath = get_file_path("请输入“导入模板.xlsx”文件路径：")
#         #     if os.path.splitext(import_template_filepath)[1].lower() != ".xlsx":
#         #         print("文件格式不正确，请选择xlsx文件")
#         #         continue
#         #     if not os.path.exists(import_template_filepath):
#         #         print("文件不存在，请检查路径")
#         #         continue
#         #     break
#         # # 复制业务号码到导入模板
#         # while True:
#         #     import_template_filepath = get_file_path("请输入“导入模板.xlsx”文件路径：")
#         #     if os.path.splitext(import_template_filepath)[1].lower() != ".xlsx":
#         #         print("文件格式不正确，请选择xlsx文件")
#         #         continue
#         #     if not os.path.exists(import_template_filepath):
#         #         print("文件不存在，请检查路径")
#         #         continue
#         #     break
#         #
#         # print(f"excel_utils 模块：{excel_utils.__file__}")  # *** 放在这里 ***
#         # 复制业务号码到导入模板
#         while True:
#             import_template_filepath = get_file_path("请输入“导入模板.xlsx”文件路径：")
#             if os.path.splitext(import_template_filepath)[1].lower() != ".xlsx":
#                 print("文件格式不正确，请选择xlsx文件")
#                 continue
#             if not os.path.exists(import_template_filepath):
#                 print("文件不存在，请检查路径")
#                 continue
#             break
#
#         print(f"excel_utils 模块：{excel_utils.__file__}")
#
#
#
#         # 调用新的函数，这个调用方式取决于你如何组织你的模块
#         # if not copy_to_template.copy_business_numbers_to_template(df, import_template_filepath):  # 调用 copy_to_template 模块中的函数
#         # # 或者直接调用修改后的函数，如果它在 business_number_utils 模块中
#         # #if not business_number_utils.copy_business_numbers_to_template(df, import_template_filepath):
#         #     print("复制业务号码失败")
#         #     exit()
#
#             #if not copy_business_numbers_to_template(df, import_template_filepath):  # 先复制数据到模板
#             #if not copy_business_numbers_to_template(df, import_template_filepath):  # 调用新的函数
#         if not copy_to_template.copy_business_numbers_to_template(df, import_template_filepath):
#             print("复制业务号码失败")
#             exit()
#             '''
#             while True:
#                 import_template_filepath = get_file_path("请输入“导入模板.xlsx”文件路径：")
#                 if os.path.splitext(import_template_filepath)[1].lower() != ".xlsx":
#                     print("文件格式不正确，请选择xlsx文件")
#                     continue
#                 if not os.path.exists(import_template_filepath):
#                     print("文件不存在，请检查路径")
#                     continue
#                 break
#             if not copy_business_numbers_to_template(df, import_template_filepath): #调用新的函数
#                 exit()
#                 '''
#             # *** 修改部分结束 ***
#
#         #insert_cols_before_vlookup = get_yn_input("数据已成功复制和筛选。是否立即在筛选后的sheet中插入新列？(y/n): ")
#         insert_cols_before_vlookup = get_yn_input("数据已成功复制。是否立即在筛选后的sheet中插入新列？(y/n): ")  # 修改提示语
#         if insert_cols_before_vlookup == 'y':
#             insert_after_sheet_name = filtered_sheet_name + "_插入后1"
#             insert_result, df = insert_columns(filtered_filepath, filtered_sheet_name, insert_after_sheet_name,
#                                                df)  # 接收df
#             if insert_result:
#                 print("筛选后的sheet新列插入成功！")
#                 df.to_excel(filtered_filepath, sheet_name=insert_after_sheet_name, index=False,
#                             engine='openpyxl')  # 保存到新的sheet
#                 try:
#                     print(f"插入新列后的DataFrame列名：{df.columns.tolist()}")
#                     filtered_sheet_name = insert_after_sheet_name
#                 except Exception as e:
#                     print(f"重新读取excel失败：{e}")
#                     exit()
#             else:
#                 print("筛选后的sheet新列插入失败！")
#                 exit()
#         else:
#             print("跳过插入新列操作。")
#
#         continue_processing = get_yn_input("是否继续处理“业务号码-LOID（SN码）”数据？(y/n): ")
#         if continue_processing == 'y':
#             while True:
#                 sn_data_filepath = get_file_path("请输入“业务号码-LOID（SN码）”数据文件路径（例如：终端数据匹配逻辑1sn码.xlsx 或 终端数据匹配逻辑1sn码.txt）：")
#                 sn_sheet = get_sheet_name("请输入“业务号码-LOID（SN码）”文件子表名称：")
#
#                 # 判断文件类型
#                 if sn_data_filepath.lower().endswith(".txt"):
#                     special_format = True
#                 elif sn_data_filepath.lower().endswith((".xls", ".xlsx")):
#                     special_format = False
#                 else:
#                     print("不支持的文件类型，请选择txt或excel文件")
#                     continue
#
#                 if "dwd_hzluheb_acc_sn_final_pg" not in os.path.basename(sn_data_filepath).lower() and "dwd_hzluheb_acc_sn_final_pg" not in sn_sheet.lower():
#                     print("文件名或子表名不包含关键字 dwd_hzluheb_acc_sn_final_pg，请检查")
#                     continue
#
#                 sorted_sn_filepath = os.path.join(os.path.dirname(sn_data_filepath), "sorted_" + os.path.basename(sn_data_filepath))
#
#                 sort_result = sort_and_save_sn_data(sn_data_filepath, sn_sheet, sorted_sn_filepath, special_format)
#                 if not sort_result:
#                     print("处理SN数据失败，请检查文件内容和格式。")
#                     continue
#
#                 try:
#                     sn_df = pd.read_excel(sorted_sn_filepath, sheet_name=sn_sheet, engine='openpyxl')
#
#                     # *** 正确的 VLOOKUP 逻辑，直接使用更新后的 df ***
#                     if "业务号码" in df.columns and "rms_access_code" in sn_df.columns and "ce_loid" in sn_df.columns and "LOID（SN码）" in df.columns:
#                         df = perform_vlookup_correct(df, sn_df)
#                         if df is not None:
#                             try:
#                                 match_sn_sheet_name = insert_after_sheet_name + "_匹配SN" # 新的 sheet 名称
#                                 with pd.ExcelWriter(filtered_filepath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
#                                     df.to_excel(writer, sheet_name=match_sn_sheet_name, index=False) # 使用新的 sheet 名称
#                                 print(f"LOID（SN码）已成功匹配并添加到文件 {filtered_filepath} 的 {match_sn_sheet_name} 工作表中。")
#                                 filtered_sheet_name = match_sn_sheet_name #更新filtered_sheet_name，方便后续操作
#
#                                 # *** 修改部分开始 ***
#                                 # 复制数据到 importCPEID2Export*.xlsx 文件
#                                 while True:
#                                     sn_export_filepath = get_file_path("请输入importCPEID2ExportSN.xlsx文件路径：")
#                                     if os.path.splitext(sn_export_filepath)[1].lower() != ".xlsx":
#                                         print("文件格式不正确，请选择xlsx文件")
#                                         continue
#                                     if not os.path.exists(sn_export_filepath):
#                                         print("文件不存在，请检查路径")
#                                         continue
#                                     break
#
#                                 while True:
#                                     mac_export_filepath = get_file_path("请输入importCPEID2ExportMAC.xlsx文件路径：")
#                                     if os.path.splitext(mac_export_filepath)[1].lower() != ".xlsx":
#                                         print("文件格式不正确，请选择xlsx文件")
#                                         continue
#                                     if not os.path.exists(mac_export_filepath):
#                                         print("文件不存在，请检查路径")
#                                         continue
#                                     break
#
#                                 if not copy_data_to_excel_with_header(df['LOID（SN码）'], sn_export_filepath, "Sheet1", "LOID"):
#                                     continue
#                                 if not copy_data_to_excel_with_header(df['ISCM终端MAC地址'], mac_export_filepath, "Sheet1", "CPEID(OUI和序列号必须填写)"):
#                                     continue
#                                 # *** 修改部分结束 ***
#
#                             except Exception as e:
#                                 print(f"保存匹配结果到目标文件时发生错误：{e}")
#                         else:
#                             print("VLOOKUP 操作失败。")
#                     else:
#                         print(f"错误：目标 DataFrame 中缺少列 '业务号码' 或 SN DataFrame 中缺少列 'rms_access_code' 或 'ce_loid' 或目标DataFrame中缺少‘LOID（SN码）’。")
#                     break #vlookup成功后退出循环
#                 except FileNotFoundError as e:
#                     print(f"文件不存在：{e}，请检查文件是否存在")
#                 except ValueError as e:
#                     print(f"工作表不存在或文件格式错误：{e}，请检查工作表是否存在")
#                 except Exception as e:
#                     print(f"其他错误：{e}")
#
#         elif continue_processing.lower() == 'n':
#             print("操作完成！")
#         else:
#             print("无效的输入，操作完成！")
#     else:
#         print("处理失败！")
#
#     print("程序结束。")
#