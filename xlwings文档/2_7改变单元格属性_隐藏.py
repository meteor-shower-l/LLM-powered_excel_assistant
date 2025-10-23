import xlwings as xw
import time
app=xw.App(add_book=False, visible=True)
WorkBook=app.books.open('test2.xlsx')
WorkSheet=WorkBook.sheets['sheet1']
time.sleep(1)
# 对行的操作
WorkSheet.range('10:10').api.Rows.Hidden = True   # 隐藏第2行
time.sleep(1)
WorkSheet.range('10:10').api.Rows.Hidden = False  # 取消隐藏第2行
time.sleep(1)

# 对列的操作：隐藏B列和C列（第2列和第3列）
WorkSheet.range('B:C').api.Columns.Hidden = True
time.sleep(1)
WorkSheet.range('B:C').api.Columns.Hidden = False
time.sleep(1)
