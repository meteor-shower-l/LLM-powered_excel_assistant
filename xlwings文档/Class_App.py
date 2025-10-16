import xlwings as xw
import time

# App类

app = xw.App(visible=None, add_book=True)
# App类构造函数用来实例化 App类对象
# visible 控制 Excel 窗口是否可见，默认为 None，即不可见
# add_book 控制启动 Excel 后是否自动创建一个空白工作簿，默认为 True，即启动 Excel 后会自动生成一个名为 Book1的空白工作簿

app.activate(steal_focus=False)
# 控制 Excel 窗口前台显示，默认为 False，温和激活窗口，不打扰用户操作

app.alert(prompt="", title=None, buttons='ok', mode=None)
# 用来弹出对话框，向用户提示信息或获取简单反馈
# prompt为对话框显示的信息（必填）  title为对话框标题  buttons为对话框按扭组合  mode为对话框图标类型

app.quit()
# 正常退出 Excel应用 无参数

app.kill()
# 强制退出 Excel 无参数






