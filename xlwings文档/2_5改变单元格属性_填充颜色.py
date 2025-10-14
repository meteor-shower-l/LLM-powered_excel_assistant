import xlwings as xw
from xlwings.utils import rgb_to_int
import time
# 初始化
app=xw.App(add_book=False,visible=True)
WorkBook=app.books.open('test1.xlsx')
WorkSheet=WorkBook.sheets['sheet1']
time.sleep(2)
# 使用.color接口，传入rgb三元组即可
# 允许对一个单元格进行操作
targetcell=WorkSheet.range('a1')
targetcell.color = (0,255,0)
time.sleep(2)
# 也允许对一个区域进行操作
targetcell=WorkSheet.range('b1:c2')
targetcell.color = (255,0,0)
time.sleep(2)