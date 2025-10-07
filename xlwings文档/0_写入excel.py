import xlwings as xw
import time

# 写入一格

# 打开应用
app=xw.App(add_book=False,visible=True)
# add_book:是否额外打开一个excel    默认值是True
# visible:是否显示excel    默认值是True
# 打开工作簿,此处是创建了一个新的工作簿
wb=app.books.add()
# 打开工作表
sht=wb.sheets['sheet1']
# 写入数据
sht.range('a1').value='kkk'
# 保存excel,由于是创建的工作簿，故需要指定名称
wb.save('test1.xlsx')
# 关闭工作簿
wb.close()
# 关闭应用
app.quit()

# 写入其他范围

app=xw.App(add_book=False)
#打开工作簿，此处是打开一个已有的工作簿
wb=app.books.open('test1.xlsx')
sht=wb.sheets['sheet1']
# 写入一行
sht.range('a2:d2').value=[1,2,3,4]
time.sleep(3)
# 写入一列
sht.range('a3').options(transpose=True).value=[5,6,7,8]
time.sleep(3)
# 插入行列(指定目标区域的左上角，注意嵌套列表即可)
sht.range('c12').value=[[1,2],[3,4]]
time.sleep(3)
wb.save()
app.quit()