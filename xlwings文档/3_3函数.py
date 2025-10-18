import xlwings as xw
import time
app=xw.App(add_book=False,visible=True)
WorkBook=app.books.open('test2.xlsx')
WorkSheet=WorkBook.sheets['sheet1']
time.sleep(2)
# 方法1：直接输入公式,向单元格输出
WorkSheet.range('A1').formula = '=SUM(B1:E1)'
time.sleep(2)
WorkSheet.range('A2').formula = '=AVERAGE(B2:E2)'
time.sleep(2)
WorkSheet.range('A3').formula = '=B3*C3'
time.sleep(2)

# 方法2：使用数组公式，向连续单元格输出
WorkSheet.range('A4:A8').formula = '=B4:B8*C4:C8'
# B4*C4,B5*C5,...
time.sleep(2)

# 数学和统计函数，求和、平均数、最大值、最小值、计数
WorkSheet.range('F1').formula = '=SUM(B1:B10)'
time.sleep(2)
WorkSheet.range('F2').formula = '=AVERAGE(B1:B10)'
time.sleep(2)
WorkSheet.range('F3').formula = '=MAX(B1:B10)'
time.sleep(2)
WorkSheet.range('F4').formula = '=MIN(B1:B10)'
time.sleep(2)
WorkSheet.range('F5').formula = '=COUNT(B1:B10)'
time.sleep(2)

# 逻辑函数,比较
WorkSheet.range('G1').formula = '=IF(B1>100,"达标","未达标")'
time.sleep(2)
WorkSheet.range('G2').formula = '=AND(B1>50,B1<200)'
time.sleep(2)
WorkSheet.range('G3').formula = '=OR(B1<0,B1>1000)'
time.sleep(2)

# 文本函数，合并、截片、取长、格式转化
WorkSheet.range('H1').formula = '=CONCATENATE("结果:",TEXT(B1,"0.00"))'  #将多个文本字符串合并为一个字符串
time.sleep(2)
WorkSheet.range('H2').formula='=TEXT(B1,"0.0")'  #  将数值转换为按指定数字格式表示的文本
time.sleep(2)
WorkSheet.range('H3').formula = '=LEFT(A1,1)'  #  从文本字符串的左侧开始提取指定数量的字符
time.sleep(2)
WorkSheet.range('H4').formula = '=LEN(A1)' # 返回文本字符串中的字符数
time.sleep(2)


# 查找和引用函数，
#  垂直查找，在表格数组的第一列中查找某个值，并返回表格数组当前行中其他列的值
WorkSheet.range('E10').formula = '=VLOOKUP(A10,A10:C12,2,FALSE)'
#=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
#要查找的值，查找范围，返回值的列号，匹配类型（FALSE=精确匹配，TRUE=近似匹配）
time.sleep(2)
#返回表格或区域中的值或值的引用
WorkSheet.range('E11').formula = '=INDEX(A10:C12,2,2)'
#=INDEX(array, row_num, column_num)
#查找范围，相对行号，相对列号
time.sleep(2)
#在单元格区域中搜索指定项，然后返回该项在区域中的相对位置
WorkSheet.range('E12').formula = '=MATCH("1.7m",B10:B12,0)'
#=MATCH(lookup_value, lookup_array, [match_type])
#要查找的值，要搜索的单行或单列区域，匹配类型（-1,0,1）
time.sleep(2)


app.quit()
