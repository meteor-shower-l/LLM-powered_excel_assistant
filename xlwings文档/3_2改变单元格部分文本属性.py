import xlwings as xw
import time
app=xw.App(add_book=False,visible=True)
WorkBook=app.books.open("test1.xlsx")
WorkSheet=WorkBook.sheets['sheet1']
# 单个字符索引（从0开始）
first_char = WorkSheet['A1'].characters[0]

# 字符切片(类似迭代器)
first_three_chars = WorkSheet['A1'].characters[0:3]  # 前3个字符


# 1. text 属性 作用: 只读指定字符范围的文本内容

# 设置单元格值
WorkSheet['A1'].value = "Hello World"
time.sleep(1)

# 获取部分文本
part_text = WorkSheet['A1'].characters[0:5].text  # "Hello"
print(part_text)
time.sleep(1)

# 修改部分文本
WorkSheet['A1'].characters[6:11].api.Text = "Excel"  # 变成 "Hello Excel"
time.sleep(1)



# 2. font 属性 作用: 设置字符的字体格式

# 设置部分字符为粗体
WorkSheet['A1'].characters[0:5].font.bold = True
time.sleep(1)

# 设置部分字符颜色
WorkSheet['A1'].characters[6:11].font.color = (255, 0, 0)  # 红色
time.sleep(1)

# 设置部分字符大小
WorkSheet['A1'].characters[0:5].font.size = 16
time.sleep(1)

# 设置斜体
WorkSheet['A1'].characters[3:7].font.italic = True
time.sleep(1)

# 关闭工作簿
WorkBook.close()
# 关闭应用
app.quit()
