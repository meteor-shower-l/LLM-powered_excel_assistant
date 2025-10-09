# 包含改变单元格行高列宽的改变方法

import xlwings as xw
import time
# 打开app
app=xw.App(add_book=False,visible=True)
# 打开工作簿
WorkBook=app.books.open("test1.xlsx")
# 打开工作表
WorkSheet=WorkBook.sheets['sheet1']
# 为了效果明显，在打开工作表，开始操作之前，先暂停了5秒（非必要操作）
time.sleep(3)

# 调整行高列宽
# 调整某一单元格列宽
TargetCell=WorkSheet.range('a1')
TargetCell.column_width=50
time.sleep(2)

# 调整某一区域列宽
TargetCell=WorkSheet.range('b1:d1')
TargetCell.column_width=50
time.sleep(2)

# 调整某一单元格行高
TargetCell=WorkSheet.range('a1')
TargetCell.row_height=100
time.sleep(2)

# 调整某一区域行高
TargetCell=WorkSheet.range('a2:a4')
TargetCell.row_height=100
time.sleep(2)

# 自动调整某一单元格行高至合适
TargetCell=WorkSheet.range('a1')
TargetCell.rows.autofit()
time.sleep(2)

# 自动调整某一区域的行高至合适
TargetCell=WorkSheet.range('a2:a4')
TargetCell.rows.autofit()
time.sleep(2)

# 自动调整某一单元格的列宽至合适
TargetCell=WorkSheet.range('a1')
TargetCell.columns.autofit()
time.sleep(2)

#自动调整某一区域的的列宽至合适
TargetCell=WorkSheet.range('b2:d2')
TargetCell.columns.autofit()
time.sleep(2)

# 自动调整某一单元格的列宽与行高
TargetCell=WorkSheet.range('a6')
TargetCell.autofit()
time.sleep(2)

# 自动调整某一区域的列宽与行高
TargetCell=WorkSheet.range('c11:d12')
TargetCell.autofit()

# 

WorkBook.close()
app.quit()
