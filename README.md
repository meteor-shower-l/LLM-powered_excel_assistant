# 表灵——基于LLM的智能Excel文件处理助手

一个基于LLM和xlwings的Excel自动化处理工具，允许用户通过自然语言指令操作Excel文件。

## 功能特点

- 通过自然语言描述Excel操作需求
- 自动解析并执行ExcelExcel操作
- 支持筛选、排序、计算、读取等常用操作
- 可视化界面，操作简单直观

## 项目结构

```
excel_assistant
│
├── Front_service.py           # 前端模块
│
├── backend.py                 # 后端模块
│
├── AI_service.py              # AI服务模块
│
├── xlwings文档                 # xlwings库部分API接口参考
│
├── APIKEY.txt                 # 存储APIKEY，用于AI读取
│
├── favicon.ico                # 程序图标文件
│
├─  dist/                      # 打包后的可执行文件目录
│   └─ 表灵.exe                 # 最终生成的GUI可执行文件
│
├── Front_service.spec         # PyInstaller打包配置文件
├── build                      # PyInstaller打包过程的临时文件
└── README.md                  # 项目说明
```

## 环境要求

- Python 3.7+
- Windows系统（xlwings需要Excel支持）
- Excel 2010或更高版本

## 通信流程

1. 前端 → AI1：发送文件路径和用户指令
2. AI1 → 前端：返回自然语言描述的操作步骤
3. 前端 → AI2：发送AI1返回的操作步骤
4. AI2 → 前端：返回后端可执行的指令格式
5. 前端 → 后端：发送AI2返回的执行指令
6. 后端 → 前端：返回操作执行结果

## 打包成EXE

可以使用pyinstaller将前端打包成exe文件：
```
pyinstaller Front_service.spec
```

注意：打包后仍需要确保AI服务和后端服务正常运行。
