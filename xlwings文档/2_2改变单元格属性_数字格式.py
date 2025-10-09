# 包含单元格的数字格式改变方法

import xlwings as xw
import time

#初始化操作对象
app=xw.App(add_book=False,visible=True)
WorkBook=app.books.open("test1.xlsx")
WorkSheet=WorkBook.sheets['sheet1']
time.sleep(2)

# 指定数字格式
targetcell=WorkSheet.range('a8')
targetcell.api.NumberFormat='0.000'    # 保留三位小数
time.sleep(2)

targetcell=WorkSheet.range('d8')
targetcell.api.NumberFormat='0.0%'    # 保留一位小数的百分比
time.sleep(2)

targetcell=WorkSheet.range('e8')
targetcell.api.NumberFormat='@'    # 文本格式
time.sleep(2)

targetcell=WorkSheet.range('f8')
targetcell.api.NUmberFormat='0.00E+00'    # 科学计数法
time.sleep(2)

WorkBook.save()
WorkBook.close()
app.quit()