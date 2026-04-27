import os
import shutil
import pandas as pd
from tkinter import Tk, filedialog

# 隐藏窗口
root = Tk()
root.withdraw()

print("=== 批量移动图片工具（自动匹配后缀，兼容带/不带后缀）===")

# 选择文件
print("请选择 Excel 文件...")
excel_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx;*.xls")])
if not excel_path:
    print("未选择文件，退出")
    exit()

print("请选择 图片所在文件夹...")
src_dir = filedialog.askdirectory(title="选择图片源文件夹")
if not src_dir:
    print("未选择文件夹，退出")
    exit()

print("请选择 移动到的目标文件夹...")
dst_dir = filedialog.askdirectory(title="选择目标文件夹")
if not dst_dir:
    print("未选择目标文件夹，退出")
    exit()

# 读取第一列（文件名，可能带后缀也可能不带）
df = pd.read_excel(excel_path)
file_names = df.iloc[:, 0].dropna().astype(str).str.strip()

# 支持的图片后缀（可自行添加）
suffix_list = ['.jpg', '.png', '.jpeg', '.bmp', '.gif', '.tiff', '.webp']
# 统一转小写，避免大小写问题
suffix_list_lower = [s.lower() for s in suffix_list]

success = 0
not_found = []

print("\n开始查找并移动文件...")

for raw_name in file_names:
    matched_path = None
    matched_suffix = None
    final_name = raw_name

    # ===================== 核心优化：自动去除已有后缀 =====================
    name_no_suffix = raw_name
    for suffix in suffix_list:
        if raw_name.lower().endswith(suffix.lower()):
            name_no_suffix = raw_name[:-len(suffix)]  # 去掉后缀
            break

    # 遍历所有后缀查找
    for suffix in suffix_list:
        test_path = os.path.join(src_dir, name_no_suffix + suffix)
        if os.path.exists(test_path):
            matched_path = test_path
            matched_suffix = suffix
            final_name = name_no_suffix
            break

    # ===================== 结束 =====================

    if matched_path:
        dst_path = os.path.join(dst_dir, final_name + matched_suffix)
        shutil.move(matched_path, dst_path)
        success += 1
        print(f"已移动：{final_name}{matched_suffix}")
    else:
        not_found.append(raw_name)

# 结果输出
print("="*50)
print(f"处理完成！")
print(f"总文件名：{len(file_names)}")
print(f"成功移动：{success}")
print(f"未找到：{len(not_found)}")

if not_found:
    print("\n未找到的文件：")
    for f in not_found[:20]:
        print(f" - {f}")

input("\n按回车键退出...")