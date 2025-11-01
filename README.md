# Excel文件处理助手

一个基于LLM和xlwings的Excel自动化处理工具，允许用户通过自然语言指令操作Excel文件。

## 功能特点

- 通过自然语言描述Excel操作需求
- 自动解析并执行ExcelExcel操作
- 支持筛选、排序、计算、读取等常用操作
- 可视化界面，操作简单直观

## 项目结构

```
excel_assistant/
│
├── frontend/                  # 前端界面
│   └── main_gui.py            # PyQt5界面实现
│
├── ai_services/               # AI服务模块
│   ├── ai1_service.py         # 指令解析服务
│   └── ai2_service.py         # 操作编码服务
│
├── backend/                   # 后端处理模块
│   ├── backend_service.py     # 后端API服务
│   └── excel_operations.py    # xlwings原子操作函数
│
├── common/                    # 公共模块
│   └── config.py              # 配置参数
│
├── assets/                    # 资源文件
│   └── icon.png               # 应用图标
│
├── requirements.txt           # 项目依赖
├── run_all.py                 # 一键启动所有服务
├── test_excel.xlsx            # 测试Excel文件
└ README.md                  # 项目说明
```

## 环境要求

- Python 3.7+
- Windows系统（xlwings需要Excel支持）
- Excel 2010或更高版本

## 安装与运行

1. 安装依赖包：
   ```
   pip install -r requirements.txt
   ```

2. 启动所有服务：
   ```
   python run_all.py
   ```

3. 使用界面输入Excel文件路径和操作指令即可

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
pyinstaller --onefile --name ExcelAssistant frontend/main_gui.py
```

注意：打包后仍需要确保AI服务和后端服务正常运行。
