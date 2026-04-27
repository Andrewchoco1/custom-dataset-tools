# 各类数据集划分脚本
## judge.py
```python
# 支持的图片格式（新增.tif，适配影像文件）
SUPPORT_FORMATS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif')
```
挺好用的一个手动评估脚本，可以手动选择数据图片文件夹，将不合格的图片筛选进入excel，便于下一步处理
## move.py
配合`judge.py`可以实现读取指定的excel，将筛选下来的图片从原有数据集中抽离出来
## Get_img_name.py
快速获取文件夹中图片的名字





