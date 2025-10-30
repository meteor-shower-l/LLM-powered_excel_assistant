import xlwings as xw
import time
import requests


def get_cmds():
    url=""
    response = requests.get(url)
    cmds_list = response.text.split(';')  # 分割指令
    for cmd in cmds_list:
        cmd=cmd.split(',')  # 分割指令的参数
    return cmds_list
#cmd 结构(handler_type,depend_id,range1,(range2),others)

# 打开app、打开具体工作表、打开工作簿
def open (path):
    app = xw.App(add_book=False, visible=True)
    workbook = app.books.open(path)
    worksheet = workbook.sheets[0]

def main():
    path=input()
    open(path)
    find_result_list=[]  #查找结果列表
    cmds_list=get_cmds()
    # 遍历指令列表
    for cmd in cmds_list:
        # cmd[1]=depend_id
        # 若不依赖于任何查找结果，则操作区域无需改变
        if(cmd[1]==0):
            handler(cmd)
        # 若依赖于查找结果，则cmd[2]=range1改变为对应的查找结果
        else:
           for result in find_result_list[cmd[1]]:
               cmd[2] =result  # 查找结果列表
               handler(cmd)

def handler(cmd):
    # 操作处理器字典
    handlers = {
        '0': handle_read,
        '1': handle_write,
        '2': handle_axis_size,
        '3': handle_autofit_axis_size,
        # 可以列举完全...
    }
    handler_type=cmd[0]
    handler=handlers[handler_type]
    if handler:
        handler(cmd)
    else:
        print(f"未知操作类型: {handler_type}")

# 读,others=none
def handle_read(cmd):
    print(worksheet.range(cmd[2]).value)

# 写,others=[axis,value]
def handle_write(cmd):
    # cmd[3]=axis,axis=0,写入行或方格；axis=1,写入列
    if cmd[3] == '0':
        worksheet.range(cmd[2]).value=cmd[4]
    elif cmd[3] == '1':
        worksheet.range(cmd[2]).options(transpose=True).value=cmd[4]
    time.sleep(T)

# 改变单元格属性_行高列宽，others=[axis,value]
def handle_axis_size(cmd):
    # cmd[3]=axis,axis=0,改变行高；axis=1,改变列宽
    if cmd[3] == '0':
        worksheet.range(cmd[2]).row_height=cmd[4]
    elif cmd[3] == '1':
        worksheet.range(cmd[2]).column_width=cmd[4]
    time.sleep(T)


# 改变单元格属性_自动行高列宽，others=[axis]
def handle_autofit_axis_size(cmd):
    # cmd[3]=axis,axis=0,改变行高；axis=1,改变列宽,axis=2,both
    if cmd[3] == '0':
        worksheet.range(cmd[2]).rows.autofit()
    elif cmd[3] ==  '1':
        worksheet.range(cmd[2]).columns.autofit()
    elif cmd[3] ==  '2':
        worksheet.range(cmd[2]).autofit()
    time.sleep(T)


# 改变单元格属性_数字格式，others=[number_formate]
def handle_number_formate(cmd):
    worksheet.range(cmd[2]).api.NumberFormate=cmd[3]
    time.sleep(T)

# 改变单元格属性_对齐方式,others=[align_way,align_id]
def handle_alignment(cmd):
    align1=[-4108,-4131,-4152,-4130,-4117]
    # 居中: -4108
    # 靠左: -4131
    # 靠右: -4152
    # 两端对齐: -4130
    # 分散对齐: -4117
    align2[-4108,-4160,-4107,-4130]
    # 居中: -4108
    # 靠上: -4160
    # 靠下: -4107
    # 两端对齐: -4130

    #cmd[3]=align_way,align_way=0,水平对齐；align_way=1,垂直对齐,align_way=2,自动换行
    if cmd[3] == '0':
        worksheet.range(cmd[2]).api.HorizontalAlignment = align1[align_id]
    elif cmd[3] == '1':
        worksheet.range(cmd[2]).api.VerticalAlignment = align2[align_id]
    elif cmd[3] == '2':
        worksheet.range(cmd[2]).WrapText = True
    time.sleep(T)

# 改变单元格属性_边框，others=[line,linestyle,weight,color]
def handle_border(cmd):
    targetcell.api.Borders(cmd[3]).LineStyle = cmd[4]
    targetcell.api.Borders(cmd[3]).Weight = cmd[5]
    targetcell.api.Borders(cmd[3]).Color = cmd[6]
    time.sleep(T)

# 改变单元格属性_颜色，others=[color]
def handle_color():
    sheet.range(cmd[2]).api.Color = cmd[3]
    time.sleep(T)

#改变单元格属性_合并，others=none
def handle_emege():
    sheet.range(cmd[2]).emerge = True
    time.sleep(T)

#改变单元格属性_拆分，others=none
def handle_unemege():
    sheet.range(cmd[2]).emerge = False
    time.sleep(T)

#改变单元格属性_隐藏，others=[axis]
def handle_hide():
    # cmd[3]=axis,axis=0,隐藏行；axis=1,隐藏列
    if cmd[3]==0:
       sheet.range(cmd[2]).api.Rows.Hidden = True
    elif cmd[3]==1:
       sheet.range(cmd[2]).api.Columns.Hidden
    time.sleep(T)

#改变单元格文本的字体，others=[name]
def handle_text_name(cmd):
   sheet.range(cmd[2]).font.name=(cmd[3])
   time.sleep(T)

#改变单元格文本的字号，others=[size]
def handle_text_size(cmd):
   sheet.range(cmd[2]).font.size=cmd[3]
   time.sleep(T)

#改变单元格文本的字体颜色，others=[color]
def handle_text_color(cmd):
   sheet.range(cmd[2]).font.color=cmd[3]
   time.sleep(T)
# 需要注意的是设置颜色时使用RGB三元组的形式传入颜色

# 改变单元格文本加粗，others=[is_bold]
def handle_text_bold(cmd):
   sheet.range(cmd[2]).font.bold=cmd[3]
   time.sleep(T)

# 改变单元格文本斜体，others=[is_italic]
def handle_text_italic(cmd):
   sheet.range(cmd[2]).font.italic=cmd[3]
   time.sleep(T)

# 改变单元格文本下划线，others=[underline_id]
def handle_text_underline(cmd):
   sheet.range(cmd[2]).api.Font.Underline=cmd[3]
   time.sleep(T)
# 4 或 True 单下划线, 5 双下划线, -4119 粗双下划线

# 改变单元格文本删除线，others=[is_strike]
def handle_text_bold(cmd):
   sheet.range(cmd[2]).api.Font.Strikethrough=cmd[3]
   time.sleep(T)

# 调用内置函数，others=[fuction]
def handle_fuction(cmd):
    worksheet.range(cmd[2]).function=cmd[3]
    time.sleep(T)
# 查找，others=[target_string]
def handle_find(target_range,target_string):
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
    # 获取第一个匹配单元格的相对地址
    first_addr = found_cell.GetAddress(False, False)

    while found_cell:
        # 将找到的单元格的相对地址加入列表
        current_addr = found_cell.GetAddress(False, False)
        found_addresses.append(current_addr)

        # 查找下一个匹配项
        found_cell = worksheet.range(target_range).api.FindNext(found_cell)

        # 如果找到下一个，但它的地址与第一个相同，说明已循环一圈，退出
        if found_cell and found_cell.GetAddress(False, False) == first_addr:
            break
    print(f'查找结果为：{found_addresses}')
    time.sleep(T)
    find_result_list.append(found_addresses)

# 排序，others=[key_list[key,order],head,orientation]
def hadle_sort(key_list,head,orientation):
    key_number=len(key_list)/2
    if key_number==1:
       worksheet.range(cmd[2]).api.Sort(
           key=key_list[0],  #排序基准列
           order=key_list[1],  #1为降序，2为升序
           Head=head,  # head:1表示有，2表示无
           Orientation =orientation  # Orientation: 1表示按行排序,2表示按列排序
       )
    elif key_number == 2:
        worksheet.range(cmd[2]).api.Sort(
            key1=key_list[0],order1=key_list[1],
            key2=key_list[2],order2=key_list[3],
            Head=head,
            Orientation =orientation
        )
    elif key_number == 3:
        worksheet.range(cmd[2]).api.Sort(
            key1=key_list[0],order1=key_list[1],
            key2=key_list[2],order2=key_list[3],
            key3=key_list[2], order3=key_list[3],
            Head=head,
            Orientation =orientation
        )
    time.sleep(T)



    # 1.得到的操作区域是单个单元格组成的列表

# 筛选，others=[field,criterial]# 列号，条件
def handle_autofilter(cmd):
    worksheet.range(cmd[2]).api.AutoFilter(Field=filed, Criteria1=criterial)
    time.sleep(T)
# 取消筛选，others=none
def handle_disautofilter(cmd):
    worksheet.api.AutoFilterMode = False  # 取消筛选
    time.sleep(T)
# 制图，others=[x_col, y_col, chart_type, data_start_row=2, chart_title]
def handle_chart(cmd):
    chart_type_list=['column_clustered','line']
    # 图表类型。'column_clustered' 为柱状图，'line'为折线图。
    # Step 1: 读取列标题和数据
    # 获取X轴和Y轴的列标题（假设在第一行）
    x_title = sheet.range(f'{cmd[2]}1').value
    y_title = sheet.range(f'{cmd[3]}1').value

    # 读取X轴和Y轴的数据（从指定行开始向下扩展）
    x_data = sheet.range(f'{cmd[2]}{data_start_row}').expand('down').value
    y_data = sheet.range(f'{cmd[3]}{data_start_row}').expand('down').value

    # 自动生成图表标题（如果未提供）
    if chart_title is None:
        chart_title = f"{y_title} vs {x_title}"

    # Step 2: 将两列数据整理到一个新的临时工作表中（确保数据连续性）
    temp_sheet = wb.sheets.add()
    # 将X轴数据写入临时工作表的A列
    temp_sheet.range('A1').value = [x_title]  # 写入X轴标题
    temp_sheet.range(f'A2').value = [[x] for x in (x_data if isinstance(x_data, list) else [x_data])]
    # 将Y轴数据写入临时工作表的B列
    temp_sheet.range('B1').value = [y_title]  # 写入Y轴标题
    temp_sheet.range(f'B2').value = [[y] for y in (y_data if isinstance(y_data, list) else [y_data])]

    # 获取临时工作表中的数据范围
    data_range = temp_sheet.range('A1').expand('table')

    # Step 3: 计算图表位置并创建图表对象
    # 将图表放在原数据表的右侧，避免覆盖[1,3](@ref)
    original_data_range = sheet.range('A1').expand()  # 假设原数据从A1开始
    chart_left = original_data_range.left + original_data_range.width + 50
    chart_top = original_data_range.top
    chart_width = 500
    chart_height = 350

    # 在原始工作表上创建图表[1](@ref)
    chart = sheet.charts.add(left=chart_left, top=chart_top, width=chart_width, height=chart_height)

    # Step 4: 设置图表数据源和类型
    chart.set_source_data(data_range)
    chart.chart_type = chart_type_list[cmd[4]]  # 设置图表类型[1](@ref)

    # Step 5: 设置图表标题和坐标轴标签[3](@ref)
    chart.api[1].SetElement(2)  # 显示标题
    chart.api[1].ChartTitle.Text = cmd[6]

    # 设置X轴标题
    chart.api[1].Axes(1).HasTitle = True
    chart.api[1].Axes(1).AxisTitle.Text = x_title

    # 设置Y轴标题
    chart.api[1].Axes(2).HasTitle = True
    chart.api[1].Axes(2).AxisTitle.Text = y_title

    # 保存工作簿
    wb.save()
    print(f"图表已成功创建！图表标题：'{cmd[6]}'")
    time.sleep(2*T)


if __name__ == "__main__":
    main()














# 2.AI返回是字符串
# 3.AI返回的字符串应该被处理为二维列表，每个第二级列表对应一个操作/.;
# 4.第二级列表的结构应为：[调用函数的类型（暂定用int），依赖于第几次查找结果（暂定int），[操作区域（是一个列表(列表元素是字符串）))，操作参数]]
# 5.
# 一人ai，一人后端，一人前端
