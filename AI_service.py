import requests
from AI_service import AI_for_coding,AI_for_divide
# 分解需求AI
# 接受用户需求与历史信息，返回分解结果
# latest_commend: 是一个字符串，记录每一次用户最新给出的要求或修改建议
# history_records: 是一个字符串列表，每个字符串记录了用户与AI的一轮对话。传入后，合并为一个字符串
# 每一个元素的格式应为:"用户:"+{用户需求}+","+"当次回答"+{AI的回答}
def AI_for_divide(latest_commend,history_records):


# 编码需求AI
# 接收经过分解的原子操作集，返回编码后的原子操作集
# 两者都是字符串
def AI_for_coding(commend):
    return encoded_commend