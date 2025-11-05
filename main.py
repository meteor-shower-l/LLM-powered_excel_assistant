import xlwings as xw
import time


class ExcelAutomation:
    def __init__(self):
        self.app = None
        self.workbook = None
        self.worksheet = None
        self.find_result_list = []  # 查找结果列表
        self.T = 1

    def open_excel(self, path):
        self.app = xw.App(add_book=False, visible=True)
        self.workbook = self.app.books.open(path)
        self.worksheet = self.workbook.sheets[0]

    def get_cmds(self, respond):
        cmds_list = respondqu.strip().split(';')  # 分割指令
        cmds_list = [
            [param.strip().strip("'") for param in cmd.split(',') if param.strip()]
            for cmd in cmds_list
            if cmd.strip()
        ]
        print(cmds_list)
        return cmds_list

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
            '6': self.handle_border,
            '7': self.handle_color,
            '8': self.handle_merge,
            '9': self.handle_unmerge,
            '10': self.handle_hide,
            '11': self.handle_text_name,
            '12': self.handle_text_size,
            '13': self.handle_text_color,
            '14': self.handle_text_bold,
            '15': self.handle_text_italic,
            '16': self.handle_text_underline,
            '17': self.handle_text_strike,
            '18': self.handle_function,
            '19': self.handle_find,
            '20': self.handle_sort,
            '21': self.handle_autofilter,
            '22': self.handle_disautofilter,
            '23': self.handle_chart,
        }
        handler_type = cmd[0]
        if handler_type in handlers:
            # 遍历ranges调用相应操作
            orign_ranges = cmd[2]
            for range in orign_ranges:
                cmd[2] = range
                handlers[handler_type](cmd)
        else:
            print(f"未知操作类型")

    def main(self, path,respond):
        self.open_excel(path)
        self.get_cmds(respond)
        # 遍历指令列表
        for cmd in cmds_list:
            # cmd[1]=depend_id
            # 若不依赖于任何查找结果，则操作区域无需改变
            if cmd[1] == '0':
                self.handler(cmd)
            # 若依赖于查找结果，则cmd[2]=range1改变为对应的查找结果
            else:
                depend_id = int(cmd[1])
                if depend_id < len(self.find_result_list):
                    cmd[2] =  self.find_result_list[depend_id]  # 查找结果列表
                    self.handler(cmd)

    # 读,others=none
    def handle_read(self, cmd):
        print(self.worksheet.range(cmd[2]).value)

    # 写,others=[axis,value]
    def handle_write(self, cmd):
        # cmd[3]=axis,axis=0,写入行或方格；axis=1,写入列
        if cmd[3] == '0':
            self.worksheet.range(cmd[2]).value = cmd[4]
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).options(transpose=True).value = cmd[4]
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
            self.worksheet.range(cmd[2]).rows.autofit()
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).columns.autofit()
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
            self.worksheet.range(cmd[2]).api.HorizontalAlignment = align1[int(cmd[4])]
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).api.VerticalAlignment = align2[int(cmd[4])]
        elif cmd[3] == '2':
            self.worksheet.range(cmd[2]).api.WrapText = True
        time.sleep(self.T)

    # 改变单元格属性_边框，others=[line,linestyle,weight,color]
    def handle_border(self, cmd):
        target_cell = self.worksheet.range(cmd[2])
        target_cell.api.Borders(int(cmd[3])).LineStyle = int(cmd[4])
        target_cell.api.Borders(int(cmd[3])).Weight = int(cmd[5])
        target_cell.api.Borders(int(cmd[3])).Color = self._parse_color(cmd[6])
        time.sleep(self.T)

    # 改变单元格属性_颜色，others=[color]
    def handle_color(self, cmd):
        self.worksheet.range(cmd[2]).color = self._parse_color(cmd[3])
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
            self.worksheet.range(cmd[2]).rows.hidden = True
        elif cmd[3] == '1':
            self.worksheet.range(cmd[2]).columns.hidden = True
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
        self.worksheet.range(cmd[2]).font.color = self._parse_color(cmd[3])
        time.sleep(self.T)

    # 改变单元格文本加粗，others=[is_bold]
    def handle_text_bold(self, cmd):
        self.worksheet.range(cmd[2]).font.bold = cmd[3].lower() == 'true'
        time.sleep(self.T)

    # 改变单元格文本斜体，others=[is_italic]
    def handle_text_italic(self, cmd):
        self.worksheet.range(cmd[2]).font.italic = cmd[3].lower() == 'true'
        time.sleep(self.T)

    # 改变单元格文本下划线，others=[underline_id]
    def handle_text_underline(self, cmd):
        underline_type=[4,5,-4119]
        self.worksheet.range(cmd[2]).api.Font.Underline = underline_type[int(cmd[3])]
        time.sleep(self.T)

    # 4 或 True 单下划线, 5 双下划线, -4119 粗双下划线

    # 改变单元格文本删除线，others=[is_strike]
    def handle_text_strike(self, cmd):
        self.worksheet.range(cmd[2]).api.Font.Strikethrough = cmd[3].lower() == 'true'
        time.sleep(self.T)

    # 调用内置函数，others=[function]
    def handle_function(self, cmd):
        self.worksheet.range(cmd[2]).formula = cmd[3]
        time.sleep(self.T)

    # 查找，others=[target_string]
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

        print(f'查找结果为：{found_addresses}')
        self.find_result_list.append(found_addresses)
        time.sleep(self.T)

    # 排序，others=[key_list[key,order]]
    def handle_sort(self, cmd):
        key_list = cmd[3].split('|')  # 假设key_list用|分隔
        head = int(cmd[4])
        orientation = int(cmd[5])
        key_number = len(key_list) // 2

        if key_number == 1:
            self.worksheet.range(cmd[2]).api.Sort(
                Key1=self.worksheet.range(key_list[0]).api,
                Order1=int(key_list[1]),
            )
        elif key_number == 2:
            self.worksheet.range(cmd[2]).api.Sort(
                Key1=self.worksheet.range(key_list[0]).api,
                Order1=int(key_list[1]),
                Key2=self.worksheet.range(key_list[2]).api,
                Order2=int(key_list[3]),
            )
        elif key_number == 3:
            self.worksheet.range(cmd[2]).api.Sort(
                Key1=self.worksheet.range(key_list[0]).api,
                Order1=int(key_list[1]),
                Key2=self.worksheet.range(key_list[2]).api,
                Order2=int(key_list[3]),
                Key3=self.worksheet.range(key_list[4]).api,
                Order3=int(key_list[5]),
            )
        time.sleep(self.T)

    # 筛选，others=[field,criteria]# 列号，条件
    def handle_autofilter(self, cmd):
        field = int(cmd[3])
        criteria = cmd[4]
        self.worksheet.range(cmd[2]).api.AutoFilter(Field=field, Criteria1=criteria)
        time.sleep(self.T)

    # 取消筛选，others=none
    def handle_disautofilter(self, cmd):
        if self.worksheet.api.AutoFilterMode:
            self.worksheet.api.AutoFilterMode = False
        time.sleep(self.T)

    # 制图，others=[x_col, y_col, chart_type, data_start_row=2, chart_title]
    def handle_chart(self, cmd):
        chart_type_list = ['column_clustered', 'line']
        x_col = cmd[3]
        y_col = cmd[4]
        chart_type_idx = int(cmd[5])
        data_start_row = int(cmd[6]) if len(cmd) > 6 else 2
        chart_title = cmd[7] if len(cmd) > 7 else None

        # 获取X轴和Y轴的列标题（假设在第一行）
        x_title = self.worksheet.range(f'{x_col}1').value
        y_title = self.worksheet.range(f'{y_col}1').value

        # 读取X轴和Y轴的数据
        x_data = self.worksheet.range(f'{x_col}{data_start_row}').expand('down').value
        y_data = self.worksheet.range(f'{y_col}{data_start_row}').expand('down').value

        # 自动生成图表标题
        if chart_title is None:
            chart_title = f"{y_title} vs {x_title}"

        # 创建临时工作表
        temp_sheet = self.workbook.sheets.add()

        # 将数据写入临时工作表
        temp_sheet.range('A1').value = [x_title]
        temp_sheet.range('A2').value = [[x] for x in (x_data if isinstance(x_data, list) else [x_data])]
        temp_sheet.range('B1').value = [y_title]
        temp_sheet.range('B2').value = [[y] for y in (y_data if isinstance(y_data, list) else [y_data])]

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

        self.workbook.save()
        print(f"图表已成功创建！图表标题：'{chart_title}'")
        time.sleep(2 * self.T)


    def set_time_interval(self, interval):
        """设置操作时间间隔"""
        self.T = interval

    def close(self):
        """关闭Excel应用"""
        if self.workbook:
            self.workbook.close()
        if self.app:
            self.app.quit()


if __name__ == "__main__":
    excel = ExcelAutomation()

