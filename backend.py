import xlwings as xw
import time


class ExcelAutomation:
    def __init__(self):
        self.app = None
        self.workbook = None
        self.worksheet = None
        self.find_result_list = []  # 查找结果列表
        self.error_list=[]  # 错误列表
        self.T = 1

    def open_excel(self, path):
        self.app = xw.App(add_book=False, visible=True)
        self.workbook = self.app.books.open(path)
        self.worksheet = self.workbook.sheets[0]
        time.sleep(3)

    def save_as(self, path):
        self.app = xw.App(add_book=False, visible=False)
        self.workbook = self.app.books.open(path)
        self.worksheet = self.workbook.sheets[0]
        self.workbook.save(r'备份.xlsx')
        self.workbook.close()
        self.app.quit()

    def close(self):
        """关闭Excel应用"""
        if self.workbook:
            self.workbook.save()
            self.workbook.close()
        if self.app:
            pass
            self.app.quit()



    def backend_main(self, path, respond):
        self.save_as(path)
        self.open_excel(path)
        cmds_list = self.get_cmds(respond)
        # 遍历指令列表
        for cmd in cmds_list:
            self.handler(cmd)
        # self.close()



    # 分割指令
    def get_cmds(self, cmd_response):
        # 按顺序分割：先分号，再逗号，最后加号
        cmds = [
            [
                [sub_item.strip() for sub_item in field.split('+') if sub_item.strip()]
                if '+' in field else field
                for field in [f.strip() for f in record.split(',') if f.strip()]
            ]
            for record in [r.strip() for r in cmd_response.split(';') if r.strip()]
        ]
        print(cmds)
        return cmds
        # cmd 结构(handler_type,depend_id,ranges[],others)

    def handler(self, cmd):
        # 操作处理器字典
        handlers = {
            '0': self.handle_read,
            '1': self.handle_write,
            '2': self.handle_axis_size,
            '3': self.handle_autofit_axis_size,
            '4': self.handle_number_formate,
            '5': self.handle_alignment,
            '6': self.handle_border_linestyle,
            '7': self.handle_border_weight,
            '8': self.handle_border_color,
            '9': self.handle_color,
            '10': self.handle_merge,
            '11': self.handle_unmerge,
            '12': self.handle_hide,
            '13': self.handle_text_name,
            '14': self.handle_text_size,
            '15': self.handle_text_color,
            '16': self.handle_text_bold,
            '17': self.handle_text_italic,
            '18': self.handle_text_underline,
            '19': self.handle_text_strike,
            '20': self.handle_find,
            '21': self.handle_sort,
            '22': self.handle_autofilter,
            '23': self.handle_deautofilter,
            '24': self.handle_chart,
        }

        handler_type = cmd[0]
        if len(cmd) <3:
            self.error_list.append(self.catch_ex(handler_type))
        else:
            # cmd[1]=depend_id
            # 若依赖于查找结果，则cmd[2]=range1改变为对应的查找结果
            if cmd[1] != '0':
                cmd[2] = self.find_result_list[int(cmd[1])-1]  # 查找结果列表
            if cmd[2] == '0':
                cmd[2]= self.get_global_range()
            if handler_type in handlers:
                # 遍历ranges调用相应操作
                orign_ranges = cmd[2]
                if type(orign_ranges) != list:
                    orign_ranges = [orign_ranges]
                for each_range in orign_ranges:
                    each_cmd = cmd.copy()
                    each_cmd[2] = each_range
                    try:
                        handlers[handler_type](each_cmd)
                    except:
                        self.error_list.append(self.catch_ex(handler_type))
            else:
                self.error_list.append(f"未知操作类型")

    def set_time_interval(self, interval):
        """设置操作时间间隔"""
        self.T = interval

    # 异常函数
    def catch_ex(self, handler_type):
        handlers_cn = {
            '0': '读取单元格内容',
            '1': '写入数据',
            '2': '调整行列尺寸',
            '3': '调整适当的行列尺寸',
            '4': '设置数字格式',
            '5': '设置对齐方式',
            '6': '设置边框线型',
            '7': '设置边框粗细',
            '8': '设置边框颜色',
            '9': '设置单元格颜色',
            '10': '合并单元格',
            '11': '取消合并单元格',
            '12': '隐藏行列',
            '13': '设置字体名称',
            '14': '设置字体大小',
            '15': '设置字体颜色',
            '16': '设置文字加粗',
            '17': '设置文字斜体',
            '18': '设置下划线',
            '19': '设置删除线',
            '20': '查找',
            '21': '数据排序',
            '22': '自动筛选',
            '23': '取消筛选',
            '24': '创建图表'
        }
        return f'{handlers_cn[handler_type]}错误'

    # 返回结果
    def get_result(self):
        if self.error_list:
            return self.error_list
        else:
            return 'success'

    # 得到所有数据的总范围
    def get_global_range(self):
        abs_addr = self.worksheet.used_range
        rel_addr = abs_addr.get_address(row_absolute=False, column_absolute=False)
        return rel_addr

    # 处理颜色
    def hex_color_to_int(self, hex_color):
        hex_color = hex_color.lstrip('#').upper()
        red = int(hex_color[0:2], 16)
        green = int(hex_color[2:4], 16)
        blue = int(hex_color[4:6], 16)
        color_int = (red << 16) | (green << 8) | blue
        return color_int

    # 将str转为bool
    def str_to_bool(self, str):
        bool_dict = {'0': False, '1': True}
        return bool_dict[str]

    #  从单元格引用中提取列范围
    def get_col_range(self,cell_ref):
        if ':' in cell_ref:
            start_cell, end_cell = cell_ref.split(':', 1)
            start_col = ''.join([c for c in start_cell if c.isalpha()])
            end_col = ''.join([c for c in end_cell if c.isalpha()])
            return f"{start_col}:{end_col}"
        else:
            col_letters = ''.join([c for c in cell_ref if c.isalpha()])
            return f"{col_letters}:{col_letters}"

    #   从单元格引用中提取行范围。
    def get_row_range(self,cell_ref):
        if ':' in cell_ref:
            start_cell, end_cell = cell_ref.split(':', 1)
            start_row = ''.join([c for c in start_cell if c.isdigit()])
            end_row = ''.join([c for c in end_cell if c.isdigit()])
            return f"{start_row}:{end_row}"
        else:
            row_digits = ''.join([c for c in cell_ref if c.isdigit()])
            return f"{row_digits}:{row_digits}"

    # 读,others=none
    def handle_read(self, cmd):
        print(self.worksheet.range(cmd[2]).value)

    # 写,others=[axis,value]
    def handle_write(self, cmd):
        # cmd[3]=axis,axis=0,写入行；axis=1,写入列;axis=2,写入方格
        if cmd[3] =='0':
            self.worksheet.range(cmd[2]).value = cmd[4]
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).options(transpose=True).value = cmd[4]
        elif cmd[3] == '2':
            self.worksheet.range(cmd[2]).value = cmd[4]
        time.sleep(self.T)

    # 改变单元格属性_行高列宽，others=[axis,value]
    def handle_axis_size(self, cmd):
        # cmd[3]=axis,axis=0,改变行高；axis=1,改变列宽
        if cmd[3] == '0':
            self.worksheet.range(cmd[2]).row_height = float(cmd[4])
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).column_width = float(cmd[4])
        time.sleep(self.T)

    # 改变单元格属性_自动行高列宽，others=[axis]
    def handle_autofit_axis_size(self, cmd):
        # cmd[3]=axis,axis=0,改变行高；axis=1,改变列宽,axis=2,both
        if cmd[3] == '0':
            row = self.get_row_range(cmd[2])
            self.worksheet.range(row).rows.autofit()
        elif cmd[3] == '1':
            column = self.get_col_range(cmd[2])
            self.worksheet.range(column).columns.autofit()
        elif cmd[3] == '2':
            self.worksheet.range(cmd[2]).autofit()
        time.sleep(self.T)

    # 改变单元格属性_数字格式，others=[number_formate]
    def handle_number_formate(self, cmd):
        self.worksheet.range(cmd[2]).api.NumberFormat = cmd[3]
        time.sleep(self.T)

    # 改变单元格属性_对齐方式,others=[align_way,align_id]
    def handle_alignment(self, cmd):
        align1 = [-4108, -4131, -4152, -4130, -4117]
        # 居中: -4108
        # 靠左: -4131
        # 靠右: -4152
        # 两端对齐: -4130
        # 分散对齐: -4117
        align2 = [-4108, -4160, -4107, -4130]
        # 居中: -4108
        # 靠上: -4160
        # 靠下: -4107
        # 两端对齐: -4130

        # cmd[3]=align_way,align_way=0,水平对齐；align_way=1,垂直对齐,align_way=2,自动换行
        if cmd[3] == '0':
            self.worksheet.range(cmd[2]).api.HorizontalAlignment = align1[int(cmd[4])-1]
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).api.VerticalAlignment = align2[int(cmd[4])-1]
        elif cmd[3] == '2':
            self.worksheet.range(cmd[2]).api.WrapText = True
        time.sleep(self.T)

    # 改变单元格属性_边框线型，others=[linestyle]
    def handle_border_linestyle(self, cmd):
        target_cell = self.worksheet.range(cmd[2])
        if cmd[3] == '0':
            cmd[3] = -4142
        for line in range(7, 11):
            target_cell.api.Borders(line).LineStyle = int(cmd[3])
        time.sleep(self.T)

    # 改变单元格属性_边框粗细，others=[weight]
    def handle_border_weight(self, cmd):
        target_cell = self.worksheet.range(cmd[2])
        for line in range(7, 11):
            target_cell.api.Borders(line).Weight = int(cmd[3])
        time.sleep(self.T)

    # 改变单元格属性_边框颜色，others=[color]
    def handle_border_color(self, cmd):
        target_cell = self.worksheet.range(cmd[2])
        color = self.hex_color_to_int(cmd[3])
        for line in range(7, 11):
            target_cell.api.Borders(line).Color = color

        time.sleep(self.T)

    # 改变单元格属性_颜色，others=[color]
    def handle_color(self, cmd):
        self.worksheet.range(cmd[2]).color = (cmd[3])
        time.sleep(self.T)

    # 改变单元格属性_合并，others=none
    def handle_merge(self, cmd):
        self.worksheet.range(cmd[2]).merge()
        time.sleep(self.T)

    # 改变单元格属性_拆分，others=none
    def handle_unmerge(self, cmd):
        self.worksheet.range(cmd[2]).unmerge()
        time.sleep(self.T)

    # 改变单元格属性_隐藏，others=[axis]
    def handle_hide(self, cmd):
        # cmd[3]=axis,axis=0,隐藏行；axis=1,隐藏列
        if cmd[3] == '0':
            row = self.get_row_range(cmd[2])
            self.worksheet.range(row).api.Rows.hidden = True
        elif cmd[3] == '1':
            column = self.get_col_range(cmd[2])
            self.worksheet.range(column).api.Columns.hidden = True
        time.sleep(self.T)

    # 改变单元格文本的字体，others=[name]
    def handle_text_name(self, cmd):
        self.worksheet.range(cmd[2]).font.name = cmd[3]
        time.sleep(self.T)

    # 改变单元格文本的字号，others=[size]
    def handle_text_size(self, cmd):
        self.worksheet.range(cmd[2]).font.size = float(cmd[3])
        time.sleep(self.T)

    # 改变单元格文本的字体颜色，others=[color]
    def handle_text_color(self, cmd):
        self.worksheet.range(cmd[2]).font.color = cmd[3]
        time.sleep(self.T)

    # 改变单元格文本加粗，others=[is_bold]
    def handle_text_bold(self, cmd):
        self.worksheet.range(cmd[2]).font.bold = self.str_to_bool(cmd[3])
        time.sleep(self.T)

    # 改变单元格文本斜体，others=[is_italic]
    def handle_text_italic(self, cmd):
        self.worksheet.range(cmd[2]).font.italic = self.str_to_bool(cmd[3])
        time.sleep(self.T)

    # 改变单元格文本下划线，others=[underline_id]
    def handle_text_underline(self, cmd):
        underline_type=[-4142, 4, 5, -4119]
        self.worksheet.range(cmd[2]).api.Font.Underline = underline_type[int(cmd[3])]
        time.sleep(self.T)
    # 4 或 True 单下划线, 5 双下划线, -4119 粗双下划线

    # 改变单元格文本删除线，others=[is_strike]
    def handle_text_strike(self, cmd):
        self.worksheet.range(cmd[2]).api.Font.Strikethrough = self.str_to_bool(cmd[3])
        time.sleep(self.T)

    # 查找,others=[target_string]
    def handle_find(self, cmd):
        target_range = cmd[2]
        target_string = cmd[3]
        first_found = self.worksheet.range(target_range).api.Find(
            What=target_string,
            LookIn=xw.constants.FindLookIn.xlValues,
            LookAt=xw.constants.LookAt.xlWhole,
            SearchOrder=xw.constants.SearchOrder.xlByRows,
            SearchDirection=xw.constants.SearchDirection.xlNext,
            MatchCase=False
        )

        if not first_found:
            self.find_result_list.append([])
            return

        found_addresses = []
        found_cell = first_found
        first_addr = found_cell.GetAddress(False, False)

        while found_cell:
            current_addr = found_cell.GetAddress(False, False)
            found_addresses.append(current_addr)
            found_cell = self.worksheet.range(target_range).api.FindNext(found_cell)
            if found_cell and found_cell.GetAddress(False, False) == first_addr:
                break
        print(found_addresses)
        self.find_result_list.append(found_addresses)
        time.sleep(self.T)

    # 排序，others=[key_list[key,order]]
    def handle_sort(self, cmd):

        sort_list = cmd[3]
        key = f'{sort_list[0]}2'
        print(key)
        order = int(sort_list[1]) + 1
        self.worksheet.range(cmd[2]).api.Sort(
            Key1=self.worksheet.range(key).api,
            Order1=order,
            Header=0,
            Orientation=1)

        time.sleep(self.T)

    # 筛选，others=[field,criteria]# 列号，条件
    def handle_autofilter(self, cmd):
        field_cell = self.worksheet.range(f'{cmd[3]}1')
        field_index = field_cell.column
        if cmd[2] == '0':
            cmd[2] = self.get_global_range()
        criteria = cmd[4]
        self.worksheet.range(cmd[2]).api.AutoFilter(Field=field_index, Criteria1=criteria)
        time.sleep(self.T)

    # 取消筛选，others=none
    def handle_deautofilter(self, cmd):
        if self.worksheet.api.AutoFilterMode:
            self.worksheet.api.AutoFilterMode = False
        time.sleep(self.T)

    # 制图，others=[x_col, y_col, chart_type, chart_title]


    def handle_chart(self, cmd):
        chart_type_list = ['column_clustered', 'line']
        x_col = cmd[3]
        y_col = cmd[4]
        chart_type_idx = int(cmd[5])
        chart_title = cmd[6]

        # 获取X轴和Y轴的列标题（假设在第一行）
        x_title = self.worksheet.range(f'{x_col}1').value
        y_title = self.worksheet.range(f'{y_col}1').value

        # 读取X轴和Y轴的数据
        x_data = self.worksheet.range(f'{x_col}2').expand('down').value
        y_data = self.worksheet.range(f'{y_col}2').expand('down').value

        # 自动生成图表标题
        if chart_title == 'None':
            chart_title = f"{y_title} vs {x_title}"

        # 创建临时工作表
        temp_sheet = self.workbook.sheets.add()

        # 将数据写入临时工作表 - 修正数据写入方式
        temp_sheet.range('A1').value = x_title
        temp_sheet.range('B1').value = y_title

        # 正确写入数据：确保X轴数据在A列，Y轴数据在B列
        for i, (x_val, y_val) in enumerate(zip(x_data, y_data), start=2):
            temp_sheet.range(f'A{i}').value = x_val
            temp_sheet.range(f'B{i}').value = y_val

        # 获取数据范围
        data_range = temp_sheet.range('A1').expand('table')

        # 计算图表位置
        original_data_range = self.worksheet.range('A1').expand()
        chart_left = original_data_range.left + original_data_range.width + 50
        chart_top = original_data_range.top
        chart_width = 500
        chart_height = 350

        # 创建图表
        chart = self.worksheet.charts.add(left=chart_left, top=chart_top, width=chart_width, height=chart_height)
        chart.set_source_data(data_range)
        chart.chart_type = chart_type_list[chart_type_idx]

        # 设置图表标题和坐标轴
        chart.api[1].SetElement(2)  # 显示标题
        chart.api[1].ChartTitle.Text = chart_title

        chart.api[1].Axes(1).HasTitle = True
        chart.api[1].Axes(1).AxisTitle.Text = x_title

        chart.api[1].Axes(2).HasTitle = True
        chart.api[1].Axes(2).AxisTitle.Text = y_title

        print(f"图表已成功创建！图表标题：'{chart_title}'")
        time.sleep(2 * self.T)




if __name__ == "__main__":
    response='''
    24,0,0,B,A,1,None;
    '''
    excel = ExcelAutomation()
    excel.backend_main(r"C:\Users\1\Desktop\学业奖学金公示名单.xlsx", response)
    print(excel.get_result())




