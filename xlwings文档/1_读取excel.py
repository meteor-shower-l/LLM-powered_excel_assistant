import xlwings as xw
import time

app=xw.App(add_book=False)
wb=app.books.open("test1.xlsx")
sht=wb.sheets['sheet1']
# 读取并输出一个单元格的内容
target_cell=sht.range('a1')
print(target_cell.value)
# 读取并输出一行的内容
target_cell=sht.range('a2:d2')
print(target_cell.value)
# 读取并输出一列的内容
target_cell=sht.range('a3:a6')
print(target_cell.value)
# 读取并输出一个范围的内容
target_cell=sht.range('c12:d13')
print(target_cell.value)
wb.close()
app.quit()
