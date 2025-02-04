import pandas as pd
import os
import pandas as pd
from openpyxl import load_workbook



def process_cpe_export_data(cpe_export_filepath):
    """处理 cpeExport 数据，插入 LOID_New 列。"""
    try:
        # 尝试使用 openpyxl 直接读取工作表
        workbook = load_workbook(cpe_export_filepath)
        sheet = workbook["sheet"]  # 替换为你的实际工作表名称
        cpe_df = pd.DataFrame(sheet.values)

        # 获取表头并设置 DataFrame 的列名
        header = [cell.value for cell in sheet[1]]  # 获取第二行作为表头 (假设表头在第二行)
        cpe_df.columns = header  # 设置 DataFrame 的列名

        # 插入 LOID_New 列 (假设 LOID 列的索引是 10，从 0 开始计数)
        cpe_df.insert(0, "LOID_New", cpe_df.iloc[:, 10])  # 使用 iloc 根据索引插入

        # 保存修改后的数据
        cpe_df.to_excel(cpe_export_filepath, sheet_name="sheet", index=False)
        return True
    except FileNotFoundError:
        print(f"错误：文件未找到：{cpe_export_filepath}")
        return False
    except KeyError:
        print(f"错误：工作表 'sheet' 未找到：{cpe_export_filepath}")  # 替换为你的实际工作表名称
        return False
    except IndexError:
        print("错误：LOID 列索引超出范围。请检查 Excel 文件。")
        return False
    except Exception as e:
        print(f"处理 cpeExport 数据出错：{e}")
        import traceback
        traceback.print_exc()
        return False
# def process_cpe_export_data(cpe_export_filepath):
#     """处理 cpeExport 数据，插入 LOID 列。"""
#     try:
#         cpe_df = pd.read_excel(cpe_export_filepath, sheet_name="sheet", engine='openpyxl')  # 读取 cpeExport 数据
#         cpe_df.insert(0, "LOID_New", cpe_df["LOID"])  # 插入 LOID 列
#         cpe_df.to_excel(cpe_export_filepath, sheet_name="sheet", index=False)  # 保存修改后的数据
#         return True
#     except Exception as e:
#         print(f"处理 cpeExport 数据出错：{e}")
#         return False
    # try:
    #     cpe_df = pd.read_excel(cpe_export_filepath, sheet_name="sheet")  # 读取 cpeExport 数据
    #     cpe_df.insert(0, "LOID_New", cpe_df["LOID"])  # 插入 LOID 列
    #     cpe_df.to_excel(cpe_export_filepath, sheet_name="sheet", index=False)  # 保存修改后的数据
    #     return True
    # except Exception as e:
    #     print(f"处理 cpeExport 数据出错：{e}")
    #     return False

def update_terminal_data(terminal_filepath, cpe_export_filepath):
    """更新终端出库报表数据，执行 VLOOKUP。"""
    try:
        # terminal_df = pd.read_excel(terminal_filepath, sheet_name="终端出库报表_筛选后1_插入后1_匹配SN")  # 读取终端出库报表数据
        # cpe_df = pd.read_excel(cpe_export_filepath, sheet_name="sheet")  # 读取 cpeExport 数据
        terminal_df = pd.read_excel(terminal_filepath, sheet_name="终端出库报表_筛选后1_插入后1_匹配SN", engine='openpyxl')  # 读取终端出库报表数据
        cpe_df = pd.read_excel(cpe_export_filepath, sheet_name="sheet", engine='openpyxl')  # 读取 cpeExport 数据


        # # 测试：打印两个 DataFrame 的信息，以便确认它们是否包含所需的列：
        # print("terminal_df:")
        # print(terminal_df.head())
        # print("cpe_df:")
        # print(cpe_df.head())

        # 执行 VLOOKUP
        terminal_df["目前在用型号2"] = terminal_df["LOID（SN码）"].apply(lambda x: get_terminal_model(x, cpe_df))

        terminal_df.to_excel(terminal_filepath, sheet_name="终端出库报表_筛选后1_插入后1_匹配SN", index=False)  # 保存修改后的数据
        return True
    except Exception as e:
        print(f"更新终端出库报表数据出错：{e}")
        import traceback
        traceback.print_exc()
        return False

def get_terminal_model(loid, cpe_df):
    """根据 LOID 查找终端型号。"""
    try:
        row = cpe_df[cpe_df["LOID"] == loid].iloc[0]  # 查找匹配的行
        return row.iloc[6]  # 返回该行第 7 列的数据 (索引为 6)
    except IndexError:  # 如果找不到匹配的行，则返回空字符串
        return ""
    except Exception as e:
        print(f"查找终端型号出错：{e}")
        return ""