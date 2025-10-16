import xlwings as xw
import time

# 创建工作表
app = xw.App(visible=True, add_book=False)
wb = app.books.add()
sht = wb.sheets[0]

# 合并 A1——B2 的单元格
sht.range('A1:B2').merge()

# 拆分 A1——B3的单元格 (若区域不可再被拆分则不执行任何操作）
sht.range('A1:B3').unmerge()


wb.save('test.xlsx')
wb.close()
app.quit()

