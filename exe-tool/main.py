import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

from processor import process_file


class ReviewTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("南方分中心舆情质检辅助工具")
        self.geometry("760x520")
        self.minsize(720, 480)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.inspector = tk.StringVar()
        self.status = tk.StringVar(value="请选择舆情质检明细 Excel 文件")

        self._build_ui()

    def _build_ui(self):
        root = tk.Frame(self, padx=18, pady=16)
        root.pack(fill="both", expand=True)

        title = tk.Label(root, text="南方分中心秒享平台舆情工单辅助审核工具", font=("Microsoft YaHei UI", 15, "bold"))
        title.pack(anchor="w", pady=(0, 12))

        form = tk.Frame(root)
        form.pack(fill="x")

        self._row(form, 0, "源文件", self.input_path, self.choose_file)
        self._row(form, 1, "输出目录", self.output_dir, self.choose_output_dir)

        tk.Label(form, text="质检人员", anchor="e", width=12).grid(row=2, column=0, sticky="e", padx=(0, 8), pady=8)
        tk.Entry(form, textvariable=self.inspector).grid(row=2, column=1, sticky="ew", pady=8)
        form.columnconfigure(1, weight=1)

        actions = tk.Frame(root)
        actions.pack(fill="x", pady=(14, 10))
        self.run_button = tk.Button(actions, text="开始处理", width=14, command=self.run)
        self.run_button.pack(side="left")
        tk.Button(actions, text="清空日志", width=10, command=lambda: self.log.delete("1.0", "end")).pack(side="left", padx=8)

        tk.Label(root, textvariable=self.status, fg="#1f5fbf", anchor="w").pack(fill="x", pady=(0, 8))

        self.log = scrolledtext.ScrolledText(root, height=18, wrap="word")
        self.log.pack(fill="both", expand=True)
        self._log("功能说明：自动复刻宏流程，包含公司筛选、超时统计、月度专项质检、舆情提醒复核、无效复核抽样和结果导出。")

    def _row(self, parent, row, label, var, command):
        tk.Label(parent, text=label, anchor="e", width=12).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=8)
        tk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=8)
        tk.Button(parent, text="选择", width=8, command=command).grid(row=row, column=2, padx=(8, 0), pady=8)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择舆情质检明细",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def run(self):
        input_path = self.input_path.get().strip()
        output_dir = self.output_dir.get().strip()
        if not input_path:
            messagebox.showwarning("提示", "请先选择源文件")
            return
        if not Path(input_path).exists():
            messagebox.showerror("错误", "源文件不存在")
            return
        if not output_dir:
            messagebox.showwarning("提示", "请选择输出目录")
            return

        self.run_button.config(state="disabled")
        self.status.set("正在处理，请稍候...")
        self._log("开始处理：" + input_path)
        threading.Thread(target=self._run_worker, args=(input_path, output_dir), daemon=True).start()

    def _run_worker(self, input_path, output_dir):
        try:
            result = process_file(input_path, output_dir, self.inspector.get())
            self.after(0, self._run_success, result)
        except Exception as exc:
            self.after(0, self._run_failed, exc)

    def _run_success(self, result):
        self.run_button.config(state="normal")
        self.status.set("处理完成")
        self._log("专项质检：" + result["special_name"])
        self._log(f"专项命中：{result['special_count']} 条")
        self._log(f"舆情提醒复核：{result['reminder_count']} 条")
        self._log(f"无效复核抽样：{result['invalid_count']} 条")
        if result["overtime_count"]:
            self._log("超时预警：" + str(result["overtime_count"]) + " 条，编号：" + "、".join(result["overtime_ids"]))
        else:
            self._log("超时预警：未查询到超时舆情")
        self._log("日报送文本：" + result["report_text"])
        self._log("输出文件：" + result["output_path"])
        messagebox.showinfo("完成", "舆情质检明细已生成。")

    def _run_failed(self, exc):
        self.run_button.config(state="normal")
        self.status.set("处理失败")
        self._log("处理失败：" + str(exc))
        messagebox.showerror("处理失败", str(exc))

    def _log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")


if __name__ == "__main__":
    app = ReviewTool()
    app.mainloop()
