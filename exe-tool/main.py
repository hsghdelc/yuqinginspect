import copy
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

from monthly_special import get_default_monthly_plans
from processor import load_config, process_file, save_config


BaseTk = TkinterDnD.Tk if TkinterDnD else tk.Tk


class ReviewTool(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("南方分中心舆情质检辅助工具")
        self.geometry("860x620")
        self.minsize(820, 560)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.inspector = tk.StringVar()
        self.status = tk.StringVar(value="请选择或拖入舆情质检明细 Excel 文件")
        self.last_report_text = ""
        self.copy_button = None

        self._build_ui()
        self._enable_drop()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=(18, 16, 18, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header,
            text="南方分中心秒享平台舆情工单辅助审核工具",
            font=("Microsoft YaHei UI", 15, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(header, text="专项配置", command=self.open_special_config).grid(row=0, column=1, padx=(10, 0))

        form = ttk.LabelFrame(self, text="文件与人员", padding=(14, 12))
        form.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, minsize=88)

        self._row(form, 0, "源文件", self.input_path, self.choose_file)
        self._row(form, 1, "输出目录", self.output_dir, self.choose_output_dir)
        ttk.Label(form, text="质检人员", width=10, anchor="e").grid(row=2, column=0, sticky="e", padx=(0, 10), pady=7)
        ttk.Entry(form, textvariable=self.inspector).grid(row=2, column=1, sticky="ew", pady=7)

        self.drop_tip = ttk.Label(
            form,
            text="可将 .xlsx / .xlsm 文件拖入窗口；如拖入不可用，请点击“选择”。",
            foreground="#546179",
        )
        self.drop_tip.grid(row=3, column=1, columnspan=2, sticky="w", pady=(4, 0))

        body = ttk.Frame(self, padding=(18, 0, 18, 0))
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        actions = ttk.Frame(body)
        actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.run_button = ttk.Button(actions, text="开始处理", command=self.run)
        self.run_button.pack(side="left")
        ttk.Button(actions, text="清空日志", command=lambda: self.log.delete("1.0", "end")).pack(side="left", padx=8)
        self.copy_button = ttk.Button(actions, text="复制日报送内容", command=self.copy_report, state="disabled")
        self.copy_button.pack(side="left")
        ttk.Label(actions, textvariable=self.status, foreground="#1f5fbf").pack(side="right")

        log_frame = ttk.LabelFrame(body, text="处理日志", padding=(8, 8))
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = scrolledtext.ScrolledText(log_frame, height=18, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew")
        self._log("使用说明：选择或拖入舆情质检明细 Excel，确认输出目录和质检人员后点击“开始处理”。处理完成后可一键复制日报送内容。")

    def _row(self, parent, row, label, var, command):
        ttk.Label(parent, text=label, width=10, anchor="e").grid(row=row, column=0, sticky="e", padx=(0, 10), pady=7)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=7)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=2, sticky="ew", padx=(10, 0), pady=7)

    def _enable_drop(self):
        if not DND_FILES:
            self.drop_tip.config(text="当前运行环境未启用拖入文件；请点击“选择”添加源文件。")
            return
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop(self, event):
        files = self.tk.splitlist(event.data)
        if not files:
            return
        path = Path(files[0])
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            messagebox.showwarning("提示", "请拖入 .xlsx 或 .xlsm 文件")
            return
        self.input_path.set(str(path))
        self.status.set("已识别拖入文件")

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

    def open_special_config(self):
        SpecialConfigWindow(self)

    def run(self):
        input_path = self.input_path.get().strip()
        output_dir = self.output_dir.get().strip()
        if not input_path:
            messagebox.showwarning("提示", "请先选择或拖入源文件")
            return
        if not Path(input_path).exists():
            messagebox.showerror("错误", "源文件不存在")
            return
        if not output_dir:
            messagebox.showwarning("提示", "请选择输出目录")
            return

        self.last_report_text = ""
        self.copy_button.config(state="disabled")
        self.run_button.config(state="disabled")
        self.status.set("正在处理，请稍候...")
        self._log("开始处理：" + input_path)
        threading.Thread(target=self._run_worker, args=(input_path, output_dir), daemon=True).start()

    def _run_worker(self, input_path, output_dir):
        try:
            result = process_file(input_path, output_dir, self.inspector.get(), progress_callback=self._thread_log)
            self.after(0, self._run_success, result)
        except Exception as exc:
            self.after(0, self._run_failed, exc)

    def _thread_log(self, text):
        self.after(0, self._log, text)

    def _run_success(self, result):
        self.run_button.config(state="normal")
        self.status.set("处理完成")
        self.last_report_text = result["report_text"]
        self.copy_button.config(state="normal")
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

    def copy_report(self):
        if not self.last_report_text:
            messagebox.showinfo("提示", "暂无可复制的日报送内容")
            return
        self.clipboard_clear()
        self.clipboard_append(self.last_report_text)
        self.status.set("日报送内容已复制")

    def _log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")


class SpecialConfigWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("专项质检配置")
        self.geometry("780x560")
        self.minsize(720, 500)
        self.transient(parent)

        self.config_data = load_config()
        self.plans = copy.deepcopy(self.config_data.get("monthly_special_plans") or get_default_monthly_plans())
        self.current_month = tk.StringVar(value="1")
        self.name_var = tk.StringVar()

        self._build_ui()
        self.month_list.selection_set(0)
        self.load_month("1")

    def _build_ui(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self, padding=(14, 14))
        left.grid(row=0, column=0, sticky="ns")
        ttk.Label(left, text="月份").pack(anchor="w")
        self.month_list = tk.Listbox(left, height=12, exportselection=False, width=12)
        self.month_list.pack(fill="y", expand=True, pady=(8, 0))
        for month in range(1, 13):
            self.month_list.insert("end", f"{month}月")
        self.month_list.bind("<<ListboxSelect>>", self.on_month_select)

        right = ttk.Frame(self, padding=(0, 14, 14, 14))
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(5, weight=1)

        ttk.Label(right, text="策略名称").grid(row=0, column=0, sticky="w")
        ttk.Entry(right, textvariable=self.name_var).grid(row=1, column=0, sticky="ew", pady=(6, 12))

        ttk.Label(right, text="质检策略").grid(row=2, column=0, sticky="w")
        self.strategy_text = scrolledtext.ScrolledText(right, height=5, wrap="word")
        self.strategy_text.grid(row=3, column=0, sticky="ew", pady=(6, 12))

        ttk.Label(right, text="质检关键词（用逗号、顿号或换行分隔）").grid(row=4, column=0, sticky="w")
        self.keyword_text = scrolledtext.ScrolledText(right, height=12, wrap="word")
        self.keyword_text.grid(row=5, column=0, sticky="nsew", pady=(6, 12))

        actions = ttk.Frame(right)
        actions.grid(row=6, column=0, sticky="ew")
        ttk.Button(actions, text="保存当前月份", command=self.save_current_month).pack(side="left")
        ttk.Button(actions, text="恢复当前月份默认值", command=self.reset_current_month).pack(side="left", padx=8)
        ttk.Button(actions, text="保存全部配置", command=self.save_all).pack(side="right")

    def on_month_select(self, _event=None):
        if not self.month_list.curselection():
            return
        self.save_current_month(show_message=False)
        month = str(self.month_list.curselection()[0] + 1)
        self.load_month(month)

    def load_month(self, month):
        self.current_month.set(month)
        plan = self.plans.get(month) or get_default_monthly_plans()[month]
        self.name_var.set(plan.get("name", ""))
        self.strategy_text.delete("1.0", "end")
        self.strategy_text.insert("1.0", plan.get("strategy", ""))
        self.keyword_text.delete("1.0", "end")
        self.keyword_text.insert("1.0", "，".join(plan.get("keywords") or []))

    def save_current_month(self, show_message=True):
        month = self.current_month.get()
        keywords = self._parse_keywords(self.keyword_text.get("1.0", "end"))
        self.plans[month] = {
            "name": self.name_var.get().strip() or f"{month}月专项",
            "strategy": self.strategy_text.get("1.0", "end").strip(),
            "keywords": keywords,
        }
        if show_message:
            messagebox.showinfo("已保存", f"{month}月专项配置已暂存")

    def reset_current_month(self):
        month = self.current_month.get()
        self.plans[month] = get_default_monthly_plans()[month]
        self.load_month(month)
        messagebox.showinfo("已恢复", f"{month}月专项配置已恢复默认值")

    def save_all(self):
        self.save_current_month(show_message=False)
        self.config_data["monthly_special_plans"] = self.plans
        save_config(self.config_data)
        self.parent.status.set("专项质检配置已保存")
        messagebox.showinfo("完成", "专项质检配置已保存，下一次处理立即生效。")

    @staticmethod
    def _parse_keywords(text):
        normalized = text.replace("，", ",").replace("、", ",").replace("\n", ",")
        return [item.strip() for item in normalized.split(",") if item.strip()]


if __name__ == "__main__":
    app = ReviewTool()
    app.mainloop()
