from backend import ExcelAutomation
from AI_service import AI_for_divide, AI_for_coding
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os


class ExcelOperationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("智能Excel文件助手")
        self.root.geometry("700x550")
        self.root.resizable(False, False)

        # 核心变量
        self.file_path = None
        self.history_records = []  # 用于AI调用的历史记录
        self.user_input_history = []  # 用于界面显示的用户输入历史
        self.is_first_input = True
        self.flag = False


        self.setup_ui()

    def setup_ui(self):

        # 设置统一的字体格式
        style = ttk.Style()
        style.configure(
            "TLabel",
            font=("微软雅黑", 11)
        )
        style.configure(
            "TButton",
            font=("微软雅黑", 11, "bold")
        )
        style.configure(
            "TFrame",
            font=("微软雅黑", 11)
        )


        """构建包含历史对话区的完整界面"""
        # 1. 文件选择区域（顶部，保持不变）
        file_frame = ttk.Frame(self.root, padding="10")
        file_frame.pack(fill=tk.X, padx=20, pady=10)
        # 文件选择区文本
        ttk.Label(file_frame, text="已选Excel文件：").pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(file_frame, text="未选择文件", foreground="gray")# 未选择填充灰色
        self.file_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        # 文件选择区按钮
        ttk.Button(
            file_frame,
            text="选择Excel文件",
            command=self.select_file
        ).pack(side=tk.RIGHT, padx=5)

        # 2. 历史对话区（新增，文件选择区下方）
        history_frame = ttk.Frame(self.root, padding="10")
        history_frame.pack(fill=tk.BOTH, padx=20, pady=5, expand=True)

        # 历史区标题
        ttk.Label(history_frame, text="历史输入记录", font=("Arial", 11, "bold")).pack(anchor=tk.W, padx=5)

        # 滚动条 + 只读文本框
        history_scroll = ttk.Scrollbar(history_frame)
        history_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

        self.history_text = tk.Text(
            history_frame,
            width=80,
            height=10,
            font=("Arial", 9),
            wrap=tk.WORD,
            state=tk.DISABLED,  # 设为只读
            yscrollcommand=history_scroll.set
        )
        self.history_text.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)
        history_scroll.config(command=self.history_text.yview)

        # 3. 输入区域（历史区下方）
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X, padx=20, pady=5)

        self.input_prompt = ttk.Label(
            input_frame,
            text="请输入你要对Excel文件进行的操作：",
            font=("Arial", 10)
        )
        self.input_prompt.pack(anchor=tk.W, padx=5, pady=5)

        self.input_text = tk.Text(
            input_frame,
            width=80,
            height=5,
            font=("Arial", 10),
            wrap=tk.WORD
        )
        self.input_text.pack(fill=tk.X, padx=5, pady=5)

        # 4. 提交按钮（底部）
        self.submit_btn = ttk.Button(
            self.root,
            text="提交需求",
            command=self.submit_needs,
            state=tk.DISABLED
        )
        self.submit_btn.pack(pady=10)

        # 初始化历史区提示文字
        self.update_history_display("暂无历史输入，请先选择文件并提交需求...")

    def select_file(self):
        """选择Excel文件（仅.xlsx/.xls）"""
        file_path = filedialog.askopenfilename(
            title="选择需要操作的Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]  # 限制Excel格式
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path), foreground="black")# 提取纯文件名增加用户反馈
            self.submit_btn.config(state=tk.NORMAL)# 启用提交需求按钮
            # 若已选择文件，更新历史区提示
            if self.user_input_history == []:
                self.update_history_display("文件已选择，请输入Excel操作需求...")

    def update_history_display(self, content, is_user_input=False):
        """更新历史对话区显示"""
        self.history_text.config(state=tk.NORMAL)  # 临时启用编辑模式

        if is_user_input:
            # 用户输入：添加序号、类型标识（初始需求/修改意见）
            input_type = "初始需求" if self.is_first_input else "修改意见"
            serial_num = len(self.user_input_history) + 1
            display_text = f"\n{serial_num}. {input_type}：\n{content}\n{'-'*100}\n"
            self.history_text.insert(tk.END, display_text)
            # 保存到显示历史列表
            self.user_input_history.append((input_type, content))
        else:
            # 提示文字：居中显示，灰色
            self.history_text.insert(tk.END, content, "hint")
            self.history_text.tag_configure("hint", foreground="gray", justify=tk.CENTER)

        # 设为只读 + 自动滚动到底部
        self.history_text.config(state=tk.DISABLED)
        self.history_text.see(tk.END)

    def submit_needs(self):
        # 1. 获取输入内容
        latest_commend = self.input_text.get("1.0", tk.END).strip()
        if not latest_commend:
            if self.is_first_input:
                messagebox.showwarning("警告", "Excel操作需求不能为空！")
            else:
                messagebox.showwarning("警告", "修改意见不能为空！")
            return

        # 2. 更新历史对话区（显示用户输入）
        if len(self.user_input_history) == 0:
            # 清空初始提示文字
            self.history_text.config(state=tk.NORMAL)
            self.history_text.delete("1.0", tk.END)
            self.history_text.config(state=tk.DISABLED)
        self.update_history_display(latest_commend, is_user_input=True)

        # 3. 切换提示语和按钮文本
        if self.is_first_input:
            self.is_first_input = False
            self.input_prompt.config(text="请给出修改意见（说明原分解不符合的地方或调整Excel操作需求）：")
            self.submit_btn.config(text="提交修改意见")

        # 4. 调用AI分解需求
        try:
            ai_divide = AI_for_divide(latest_commend, self.history_records)
            print(ai_divide)
        except Exception as e:
            messagebox.showerror("错误", f"AI分解异常：{str(e)}")
            return

        if not ai_divide:
            messagebox.showwarning("警告", "AI分解失败，请重新补充Excel操作需求或修改意见！")
            self.history_records.append(f"用户:{latest_commend},当次回答:AI分解失败")
            self.input_text.delete("1.0", tk.END)
            return

        # 5. 确认分解结果
        user_confirm = messagebox.askyesno(
            title="确认Excel操作分解",
            message=f"AI 已将你的需求分解为以下原子操作：\n\n{ai_divide}\n\n是否符合你的预期？"
        )

        if user_confirm:
            # 6. 执行编码和后端操作
            self.flag = True
            messagebox.showinfo("提示", "请点击确定以执行下一步操作！")

            try:
                encoded_commend = AI_for_coding(ai_divide)
                print(encoded_commend)
            except Exception as e:
                messagebox.showerror("错误", f"AI编码异常：{str(e)}")
                self.flag = False
                return

            if not encoded_commend:
                messagebox.showerror("错误", "AI编码失败，无法执行Excel操作！")
                self.flag = False
                self.input_text.delete("1.0", tk.END)
                return

            try:
                excel=ExcelAutomation()
                excel.backend_main(self.file_path, encoded_commend)
                backend_result = excel.get_result()
                messagebox.showinfo(
                    "Excel操作成功",
                    f"Excel操作执行完成！\n\n后端返回结果：\n{backend_result}"
                )
                self.root.quit()
            except Exception as e:
                messagebox.showerror(
                    "Excel执行失败",
                    f"后端执行Excel操作失败：\n{str(e)}\n\n请修改意见后重新尝试"
                )
                self.flag = False
                self.input_text.delete("1.0", tk.END)

        else:
            # 7. 记录历史并继续
            self.flag = False
            self.history_records.append(f"用户:{latest_commend},当次回答:{ai_divide}")
            messagebox.showinfo(
                "提示",
                "已记录本次Excel需求和分解结果，将用于优化下次AI分解\n请补充修改意见后重新提交"
            )
            self.input_text.delete("1.0", tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelOperationGUI(root)
    root.mainloop()
    root.destroy()