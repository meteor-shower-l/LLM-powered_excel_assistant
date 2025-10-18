import xlwings as xw

app=xw.App(add_book=False,visible=True)
WorkBook = app.books.open('test1.xlsx')
WorkSheet = WorkBook.sheets['sheet1']

targetcell = WorkSheet.range('J2:L5')
targetcell.api.Sort(
    Key1 = WorkSheet.range('L2').api, # 指定L为排序基准列
    Order1 = 2,# 指定为升序
    Key2 = WorkSheet.range('K2').api, # 指定K为排序第二基准列（在排序基准列相同时，依次为比较依据)
    Order2 = 1, # 指定为降序
    Orientation = 1, # Orientation: 1表示按行排序,2表示按列排序
)
# 相较于pandas排序，使用xlwings库内置排序的优点是允许携带其他区域一起变，缺点是难以实现十分复杂的程序
# pandas的优势在于可以处理较为复杂的比较逻辑，缺点是代码逻辑区别较大
# 本处文档说明了内置方法的使用，具体排序算法的使用，留待课上讨论决定使用内置还是pandas