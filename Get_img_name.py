import os
import pandas as pd
from tkinter import Tk, filedialog

# 隐藏窗口
root = Tk()
root.withdraw()

print("=== 读取图片名称（不带后缀）并导出到 Excel ===")

# 选择图片文件夹
print("请选择 图片所在文件夹...")
img_dir = filedialog.askdirectory(title="选择图片文件夹")
if not img_dir:
    print("未选择文件夹，退出")
    exit()

# 支持的图片格式
img_suffix = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp']

# 读取所有图片（去掉后缀）
img_names = []
for file in os.listdir(img_dir):
    # 只处理图片
    if any(file.lower().endswith(suf) for suf in img_suffix):
        # 去掉后缀 🔥
        name_without_suffix = os.path.splitext(file)[0]
        img_names.append(name_without_suffix)

# 写入 Excel（第一列）
df = pd.DataFrame(img_names, columns=['文件名'])
excel_save_path = os.path.join(img_dir, '图片名称列表.xlsx')
df.to_excel(excel_save_path, index=False)

# 完成提示
print("="*50)
print(f"✅ 处理完成！")
print(f"共读取：{len(img_names)} 个文件")
print(f"已生成 Excel：{excel_save_path}")
print(f"✅ 文件名**不带后缀**，已写入第一列")