import xlwings as xw
import time

app=xw.App(add_book=False,visible=True)

work_book=app.books.open('test1.xlsx')

work_sheet=work_book.sheets['sheet1']
time.sleep(2)

# 水平对齐

# 单个单元格水平左对齐
targetcell=work_sheet.range('a2')
targetcell.api.HorizontalAlignment=-4131
time.sleep(2)

# 单个单元格水平居中对齐
targetcell=work_sheet.range('a3')
targetcell.api.HorizontalAlignment=-4108
time.sleep(2)

# 单个单元格水平右对齐
targetcell=work_sheet.range('a4')
targetcell.api.HorizontalAlignment=-4152
time.sleep(2)

# -4131: 水平方向靠左
# -4108：水平方向居中
# -4152：水平方向靠右

# 垂直对齐

# 单个单元格垂直靠上
targetcell=work_sheet.range('a8')
targetcell.api.VerticalAlignment=-4160
time.sleep(2)

# 单个单元格垂直居中
targetcell=work_sheet.range('b8')
targetcell.api.VerticalAlignment=-4108
time.sleep(2)

# 单个单元格垂直靠下
targetcell=work_sheet.range('c8')
targetcell.api.VerticalAlignment=-4107
time.sleep(2)

# -4160：垂直方向靠上
# -4108：垂直方向居中
# -4107：垂直方向靠下

# 以上方法对区域也适用，直接推广即可