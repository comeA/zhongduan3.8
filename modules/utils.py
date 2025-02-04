# Excel_处理终端数据 V3.8/modules/utils.py
import os

def get_sheet_name(prompt):
    """循环提示用户输入工作表名称，直到输入非空值为止。"""
    while True:
        sheet_name = input(prompt).strip()
        if sheet_name:
            return sheet_name
        else:
            print("工作表名称不能为空，请重新输入。")

def get_file_path(prompt, check_exists=True):
    """循环提示用户输入文件路径，直到满足条件为止。
    prompt: 提示用户的字符串。
    check_exists: 是否检查文件路径是否存在，默认为 True。
    """
    while True:
        file_path = input(prompt).replace("\\", "/").strip() # 统一使用正斜杠
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
    """循环提示用户输入 y/n，直到输入正确为止。"""
    while True:
        user_input = input(prompt).strip().lower()
        if user_input in ('y', 'n'):
            return user_input
        else:
            print("请输入 y 或 n。")

def create_directory(path):
    """创建目录，如果目录已存在则不进行任何操作。"""
    try:
        os.makedirs(path, exist_ok=True)  # exist_ok=True 防止目录已存在时抛出异常
        print(f"成功创建目录：{path}")
        return True
    except OSError as e:
        print(f"创建目录 {path} 失败: {e}")
        return False

def check_file_exists(file_path):
    """检查文件是否存在"""
    if os.path.exists(file_path):
        return True
    else:
        print(f"文件路径 {file_path} 不存在，请检查路径")
        return False

def check_excel_file(file_path):
    """检查文件是否为excel文件"""
    if os.path.splitext(file_path)[1].lower() != ".xlsx":
        print("文件格式不正确，请选择xlsx文件")
        return False
    return True

def get_yes_no_input(prompt):
    """循环提示用户输入 y/n，直到输入正确为止。"""
    while True:
        user_input = input(prompt).strip().lower()
        if user_input in ('y', 'n'):
            return user_input
        else:
            print("请输入 y 或 n。")