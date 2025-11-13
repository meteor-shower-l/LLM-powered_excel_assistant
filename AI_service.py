import requests
api_key ='JsYQOBMSboleXdNueMto:SDdlAAOGeHvPfsedsPYT'
url = 'https://spark-api-open.xf-yun.com/v1/chat/completions'
# 用于将列表整合为一整个字符串
def integrate_history_records(history_records):
    history_records_str =""
    for item in history_records:
        flag =1
        history_records_str += f"第{flag}轮对话：\n"+item
        flag += 1
    return history_records_str
# 获得AI回答
def get_answer(message):
    #初始化请求体
    headers = {
        'Authorization':api_key,
        'content-type': "application/json"
    }
    body = {
        "model": "4.0Ultra",
        "user": "user",
        "messages":message,
        "stream": False,
        "tools": [
            {
                "type": "web_search",
                "web_search": {
                    "enable": False,
                    "search_mode":"deep"
                }
            }
        ]
    }
    response = requests.post(url=url,json= body,headers= headers)
    if('error' in response.json()):
        return "运行错误"
    else:
        final_response=response.json()['choices'][0]['message']['content']
        return final_response
prompt_AI_for_divide_divider = """
## 角色
你是一个Excel操作专家，擅长根据用户明确的指令精准执行各类Excel自动化任务，严格遵守原子化指令规范进行数据交互与格式调整。
## 技能
1. **读取和写入单元格内容**：
  - 根据用户提供的行列号或查找结果，准确读取指定单元格的内容。
  - 按照用户的指令，将特定内容写入到指定的单元格中。
2. **调整行宽和行高**：
  - 能够依据用户需求，精确设置表格中某一行或多行的宽度和高度数值。
3. **自动调整行高列宽**：
  - 自动适应单元格内文本或其他内容的长度，优化行高和列宽以完整展示信息。
4. **调整单元格数字格式**：
  - 可将选定单元格内的数据转换为常规、日期、小数点后几位等不同的数字显示格式。
5. **调整单元格对齐方式**：
  - 设定单元格内文本的水平（左、中、右）和垂直（上、中、下）对齐模式。
6. **调整单元格边框**：
  - 为指定单元格改变不同样式（如实线、虚线）、颜色和粗细的边框线条。
7. **调整单元格颜色**：
  - 改变单元格的背景填充色。
8. **合并与拆分单元格**：
  - 把多个相邻且符合要求的单元格合并成一个大的单元格区域；反之，也能将已合并的单元格重新拆分还原。
9. **隐藏单元格**：
  - 根据需要隐藏特定的行、列或者单个单元格。
10. **调整单元格文本字体相关属性**：
    - 包括修改字体类型、字号大小、加粗、斜体、下划线、删除线以及文本颜色等多方面的文字样式设定。
11. **查找功能**：
    - 在工作表范围内依照给定的关键字符或者其他条件快速定位目标单元格位置，并记录每次查找的结果以便后续引用。
12. **排序操作**：
    - 针对一行、一列甚至整个数据区域内的信息按照升序或者降序规则重新排列顺序。
13. **筛选功能启用与关闭**：
    - 隐藏不满足筛选条件的行或列
14. **制图能力**：
    - 基于已有的数据源创建各种类型的图表用于直观呈现数据之间的关系变化趋势等情况。
## 限制
- 所有操作必须严格限定在用户明确指定的范围内执行，禁止越界修改未授权区域。
- 仅支持以下原子化指令格式：“单元格”+“操作”+“操作参数”（例：将A1到A3改变单元格颜色为黄色）；“目标为第几次查找的结果”+“操作”+“操作参数”（例：将第二次查找到的范围增加字号至18）。需要注意的是，你需要声明的是第几次查找而非第几步操作。
- 你对单元格的描述只有两种方法，一种是提供行列号，另一种是指出是第几次查找的结果。例如你不能说“包含小明的单元格”或“值大于5的单元格”，对于此类操作，你必须通过查找实现。
- 禁止提供任何建议、解释或额外信息，仅输出分解得到的原子操作。
- 必须准确区分不同查找次数的目标范围，避免混淆操作对象。
- 若用户给出的不是完整需求而是纠正，必须结合历史信息，给出完整的原子操作集
"""
# 分解需求AI
# 接受用户需求与历史信息，返回分解结果
# latest_commend: 是一个字符串，记录每一次用户最新给出的要求或修改建议
# history_records: 是一个字符串列表，每个字符串记录了用户与AI的一轮对话。传入后，合并为一个字符串
# 每一个元素的格式应为:"用户:"+{用户需求}+","+"当次回答"+{AI的回答}
def AI_for_divide(latest_commend,history_records):
    history_records_str = integrate_history_records(history_records)
    message = (prompt_AI_for_divide_divider 
    +f'请将用户的需求{latest_commend}分解为各个原子操作的组合，'
    +'以下是历史对话:\n请结合历史对话进行分解'+history_records_str)
    response = get_answer(message)

# 编码需求AI
# 接收经过分解的原子操作集，返回编码后的原子操作集
# 两者都是字符串
def AI_for_coding(commend):
    return encoded_commend
