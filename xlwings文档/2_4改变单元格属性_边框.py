import xlwings as xw
from xlwings.utils import rgb_to_int
import time
# 初始化
app=xw.App(add_book=False,visible=True)
WorkBook=app.books.open('test1.xlsx')
WorkSheet=WorkBook.sheets['sheet1']
time.sleep(2)

#
targetcell=WorkSheet.range('b2')

targetcell.api.Borders(7).LineStyle = 2  # 左边框，虚线
targetcell.api.Borders(7).Weight = 3     # 左边框，粗细
targetcell.api.Borders(7).Color = rgb_to_int((255,0,0)) # 左边框，红色
time.sleep(2)

targetcell.api.Borders(8).LineStyle = 3  # 上边框，点线
targetcell.api.Borders(8).Weight = 3     # 上边框，粗细
targetcell.api.Borders(8).Color = rgb_to_int((0,255,0)) # 上边框，绿色
time.sleep(2)

targetcell.api.Borders(9).LineStyle = 4  # 下边框，点划线
targetcell.api.Borders(9).Weight = 3     # 下边框，粗细
targetcell.api.Borders(9).Color = rgb_to_int((0,0,255)) # 下边框,蓝色
time.sleep(2)

targetcell.api.Borders(10).LineStyle = 5 # 右边框，双点划线
targetcell.api.Borders(10).Weight = 3    # 右边框，粗细
targetcell.api.Borders(10).Color = rgb_to_int((128,128,128)) # 右边框,不知道什么色
time.sleep(2)

# 设置内部边框（单元格之间的线）
# 此处似乎没用
targetcell.api.Borders(11).LineStyle = 2  # 内部垂直边线，虚线
targetcell.api.Borders(12).LineStyle = 2  # 内部水平边线，虚线
targetcell.api.Borders(12).Color = rgb_to_int((255,0,0))
time.sleep(2)

# ​​边框索引​​:
# 5: 单元格从左上到右下的对角线
# 6: 单元格从左下到右上的对角线
# 7: 左边框
# 8: 上边框
# 9: 下边框
# 10: 右边框
# 11: 内部垂直边线
# 12: 内部水平边线

# 线型索引:.LineStyle
# 1: 实线
# 2: 虚线
# 3: 点线
# 4: 点划线
# 5: 双点划线
# -4142: 无边框

# 粗细索引；.Weight
# 1: 细线
# 2: 细 
# 3: 中
# 4: 粗

# 颜色：.Color
# 需要使用0x开头的16进制数
# 但是可以import xlwings.utils 的 rgb_to_int
# 来将三元组转换为16进制数