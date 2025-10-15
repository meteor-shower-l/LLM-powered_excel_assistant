# 字体，字号，颜色，加粗，斜体，下划线，删除线，上下标的设置

import xlwings as xw
import time

app=xw.App(add_book=False,visible=True)
WorkBook=app.books.open("test1.xlsx")
WorkSheet=WorkBook.sheets['sheet1']
time.sleep(2)

# 改变一个单元格的字体
targetcell=WorkSheet.range('a1')
targetcell.font.name="微软雅黑"
time.sleep(2)

# 改变一个区域的字体
targetcell=WorkSheet.range('a2:a4')
targetcell.font.name=('Times New Roman')
time.sleep(2)

# 改变一个单元格的字号
targetcell=WorkSheet.range('a1')
targetcell.font.size=20
time.sleep(2)

# 改变一个区域的字号
targetcell=WorkSheet.range('a2:a4')
targetcell.font.size=20
time.sleep(2)

# 需要注意的是设置颜色时使用RGB三元组的形式传入颜色
# 设置一个单元格的字体颜色
targetcell=WorkSheet.range('a1')
targetcell.font.color=(255,0,0)
time.sleep(2)

# 设置一个区域的字体颜色
targetcell=WorkSheet.range('a2:a4')
targetcell.font.color=(0,0,255)
time.sleep(2)

# 设置一个单元格的加粗
targetcell=WorkSheet.range('a1')
targetcell.font.bold=True
time.sleep(2)

# 设置一个区域的加粗
targetcell=WorkSheet.range('b2:d2')
targetcell.font.bold=True
time.sleep(2)

# 设置一个单元格的斜体
targetcell=WorkSheet.range('a1')
targetcell.font.italic=True
time.sleep(2)

# 设置一个区域的斜体
targetcell=WorkSheet.range("b2:d2")
targetcell.font.italic=True
time.sleep(2)

# 需要注意的是，下划线,删除线属性是更加底层的属性，所以需要调用.api来访问VBA对象模型方法,因此写法与其他不同
# VBA是微软为excel设计的一套语言，xlwings库可以与VBA协作
# 另外，使用.api时，其后接的属性或方法​​首字母通常需要大写
# 关于.api.Font.Underline的参数:
# 4 或 True 单下划线, 5 双下划线, -4119 粗双下划线

# 设置一个单元格的单下划线
targetcell=WorkSheet.range('a1')
targetcell.api.Font.Underline=4
time.sleep(2)

# 设置一个区域的双下划线
targetcell=WorkSheet.range('b2:d2')
targetcell.api.Font.Underline=5
time.sleep(2)

# 设置一个单元格的删除线
targetcell=WorkSheet.range('a1')
targetcell.api.Font.Strikethrough=True
time.sleep(2)

# 设置一个区域的删除线
targetcell=WorkSheet.range('b2:d2')
targetcell.api.Font.Strikethrough=True
time.sleep(2)

# 在设置上下标中，也需要使用.api来访问VBA的更加底层的属性
# GetCharater(a,b)的含义是：从第a个字符开始，对b个字符执行上标或下标操作。其中一定要注意的是本处从1开始计。
# 再次强调，从1开始计！！！例如：E=mc2中，E是第1个，=是第2个，m是第3个
# 另外需要注意的是：
# 上下标操作不能对区域施加，正如下边代码所见的，这是对具体文本的精细化操作，只能对单元格进行操作

# 设置上标
targetcell=WorkSheet.range('b1')
targetcell.value='E=mc2'
time.sleep(2)
# 为了强调是两步操作（先指定内容，再指定上下标），此处特地暂停
targetcell.api.GetCharacters(5,1).Font.Superscript=True
time.sleep(2)

# 设置下标
targetcell=WorkSheet.range('c1')
targetcell.value ='H2O'
time.sleep(2)
# 为了强调是两步操作（先指定内容，再指定上下标），此处特地暂停
targetcell.api.GetCharacters(2,1).Font.Subscript = True
time.sleep(2)

WorkBook.save()
WorkBook.close()
app.quit()