import xlwings as xw
def xlwings_find_all(start_cell,end_cell,target_string):
    target_range = f'{start_cell}:{end_cell}'
    first_found = WorkSheet.range(target_range).api.Find(
        What=target_string,
        # 指定搜索字符串
        LookIn = xw.constants.FindLookIn.xlValues,
        # 使用LookIn参数指定是在单元格的值、公式还是批注中查找：xlValues为值, xlFormulas为公式, xlComments为批注
        LookAt=xw.constants.LookAt.xlWhole,
        # 使用LookAt参数控制全字匹配或部分匹配:xlWhole为完全匹配，xlPart为部分匹配
        SearchOrder = xw.constants.SearchOrder.xlByRows,
        # 使用SearchOrder参数指定搜索顺序:xlByRows为按行,xlByColumns为按列
        SearchDirection = xw.constants.SearchDirection.xlNext,
        # 使用SearchOrder参数指定搜索方向:xlNext为向后搜索，xlPrevious为向前搜索
        MatchCase =False,
        # 使用MatchCase参数指定搜索是否区分大小写
    )
    if not first_found:
        return []
    found_addresses = []
    found_cell = first_found
    first_addr = found_cell.Address
    while found_cell:
        # 将找到的单元格加入列表
        found_addresses.append(found_cell)
        # 查找下一个匹配项
        found_cell = WorkSheet.range(target_range).api.FindNext(found_cell)
        # 如果找到下一个，但它的地址与第一个相同，说明已循环一圈，退出
        if found_cell and found_cell.Address == first_addr:
            break
    return found_addresses

app = xw.App(add_book = False,visible = True)
WorkBook = app.books.open('test1.xlsx')
WorkSheet = WorkBook.sheets['sheet1']
# 测试代码
result = xlwings_find_all('G6','J8',1)
print(result[0].value)