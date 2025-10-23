import xlwings as xw
import time

# 打开app、打开具体工作表、打开工作簿
def main(ai_returned_list):
    # 用于储存查找函数的结果
    serach_result_list=[[]] # 由于ai给定的是“第i次查找的结果”，所以search_result_list[0]定义问问空列表
    # 遍历操作队列
    for i in ai_returned_list:
        # 若不依赖于任何查找结果，则操作区域无需改变
        if(i[1]==0):
            pass
        # 若依赖于查找结果，则操作区域改变位对应的查找结果
        else:
            i[2][0] =serach_result_list[i[1]]
        use(i[0],i[1],i[2])
        # 需要注意的是，在find函数最后，需要把结果添加到search_result_list

# 1.得到的操作区域是单个单元格组成的列表
# 2.AI返回是字符串
# 3.AI返回的字符串应该被处理为二维列表，每个第二级列表对应一个操作/.;
# 4.第二级列表的结构应为：[调用函数的类型（暂定用int），依赖于第几次查找结果（暂定int），[操作区域（是一个列表(列表元素是字符串）))，操作参数]]
# 5.
# 一人ai，一人后端，一人前端
def use(function_type,function_depend,function_param_list):
    # function_type :指定函数类别
    # function_depend :指定区域依赖于哪个查找结果
    # function_param_list :获得ai提供的参数
    # function_param_list :[targetcell,other_param_list]
    pass
    if:
    else