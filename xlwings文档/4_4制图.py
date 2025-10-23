import xlwings as xw


def create_chart_from_columns(excel_path, x_col, y_col, chart_type='column_clustered', sheet_name=None,
                              data_start_row=2, chart_title=None):
    """
    根据Excel中任意指定的两列数据创建图表（柱状图或折线图）

    Parameters:
    excel_path (str): Excel文件路径。
    x_col (str): 作为X轴的列字母，例如 'A'。
    y_col (str): 作为Y轴的列字母，例如 'B'。
    chart_type (str): 图表类型。'column_clustered'为柱状图，'line'为折线图。
    sheet_name (str): 工作表名称，默认为第一个工作表。
    data_start_row (int): 数据起始行号（从1开始），通常第1行为标题，数据从第2行开始。
    chart_title (str): 图表主标题。如果为None，则自动生成。
    """

    # 启动Excel应用程序（后台运行）
    app = xw.App(visible=False, add_book=False)


    # 打开工作簿并选择工作表
    wb = app.books.open(excel_path)
    if sheet_name:
        sheet = wb.sheets[sheet_name]
    else:
        sheet = wb.sheets[0]

    # Step 1: 读取列标题和数据
    # 获取X轴和Y轴的列标题（假设在第一行）
    x_title = sheet.range(f'{x_col}1').value
    y_title = sheet.range(f'{y_col}1').value

    # 读取X轴和Y轴的数据（从指定行开始向下扩展）
    x_data = sheet.range(f'{x_col}{data_start_row}').expand('down').value
    y_data = sheet.range(f'{y_col}{data_start_row}').expand('down').value

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
    chart.chart_type = chart_type  # 设置图表类型[1](@ref)

    # Step 5: 设置图表标题和坐标轴标签[3](@ref)
    chart.api[1].SetElement(2)  # 显示标题
    chart.api[1].ChartTitle.Text = chart_title

    # 设置X轴标题
    chart.api[1].Axes(1).HasTitle = True
    chart.api[1].Axes(1).AxisTitle.Text = x_title

    # 设置Y轴标题
    chart.api[1].Axes(2).HasTitle = True
    chart.api[1].Axes(2).AxisTitle.Text = y_title

    # 保存工作簿
    wb.save()
    print(f"图表已成功创建！图表标题：'{chart_title}'")

    wb.close()
    app.quit()

create_chart_from_columns(excel_path=r"C:\Users\48994\Desktop\M_time_movies.xlsx",
                          x_col='A', y_col='D', chart_type='column_clustered',
                          chart_title='电影评分')
