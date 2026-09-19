import copy
import json
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageDraw

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

from monthly_special import get_default_monthly_plans
from processor import (
    DEFAULT_SCHEME,
    DEFAULT_SCHEME_ID,
    copy_json,
    export_quality_file,
    load_config,
    load_workbook_headers,
    process_file,
    save_config,
    validate_config_for_file,
)


BaseTk = TkinterDnD.Tk if TkinterDnD else tk.Tk

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BG = "#f1f6fa"
SURFACE = "#ffffff"
TEXT = "#102a43"
MUTED = "#73839a"
TEAL = "#07999d"
TEAL_HOVER = "#07878b"
BORDER = "#d7e2ec"
FONT = "Microsoft YaHei UI"
ICON_CACHE = {}


def ui_icon(name, color="#1766a3", size=18):
    key = (name, color, size)
    if key in ICON_CACHE:
        return ICON_CACHE[key]
    scale = 3
    width = size * scale
    image = Image.new("RGBA", (width, width), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    stroke = scale * 2
    pad = 4 * scale
    mid = width // 2

    if name in {"settings", "special"}:
        draw.ellipse((pad, pad, width - pad, width - pad), outline=color, width=stroke)
        draw.ellipse((mid - 3 * scale, mid - 3 * scale, mid + 3 * scale, mid + 3 * scale), outline=color, width=stroke)
        draw.line((mid, pad - scale, mid, pad + 4 * scale), fill=color, width=stroke)
        draw.line((mid, width - pad + scale, mid, width - pad - 4 * scale), fill=color, width=stroke)
    elif name in {"import", "file"}:
        draw.rounded_rectangle((pad + 2 * scale, pad, width - pad - 2 * scale, width - pad), radius=2 * scale, outline=color, width=stroke)
        draw.line((mid, pad + 4 * scale, mid, width - pad - 6 * scale), fill=color, width=stroke)
        draw.line((mid - 4 * scale, mid, mid, mid + 4 * scale, mid + 4 * scale, mid), fill=color, width=stroke)
    elif name in {"export", "download"}:
        draw.line((mid, pad, mid, width - pad - 6 * scale), fill=color, width=stroke)
        draw.line((mid - 5 * scale, mid, mid, mid + 5 * scale, mid + 5 * scale, mid), fill=color, width=stroke)
        draw.line((pad, width - pad, width - pad, width - pad), fill=color, width=stroke)
    elif name == "restore":
        draw.arc((pad, pad, width - pad, width - pad), 40, 320, fill=color, width=stroke)
        draw.polygon(((pad, mid), (pad + 6 * scale, mid - 4 * scale), (pad + 6 * scale, mid + 4 * scale)), fill=color)
    elif name in {"task", "monthly", "log"}:
        draw.rounded_rectangle((pad + 2 * scale, pad + 3 * scale, width - pad - 2 * scale, width - pad), radius=2 * scale, outline=color, width=stroke)
        draw.rounded_rectangle((mid - 4 * scale, pad, mid + 4 * scale, pad + 5 * scale), radius=2 * scale, fill=color)
        for y in (mid - 2 * scale, mid + 4 * scale):
            draw.line((pad + 6 * scale, y, width - pad - 6 * scale, y), fill=color, width=stroke)
    elif name == "play":
        draw.polygon(((pad + 4 * scale, pad), (width - pad, mid), (pad + 4 * scale, width - pad)), fill=color)
    elif name == "trash":
        draw.rounded_rectangle((pad + 4 * scale, pad + 6 * scale, width - pad - 4 * scale, width - pad), radius=2 * scale, outline=color, width=stroke)
        draw.line((pad + 2 * scale, pad + 4 * scale, width - pad - 2 * scale, pad + 4 * scale), fill=color, width=stroke)
        draw.line((mid - 3 * scale, pad, mid + 3 * scale, pad), fill=color, width=stroke)
    elif name == "chart":
        for idx, bar_height in enumerate((7, 11, 15)):
            x = pad + idx * 6 * scale
            draw.rounded_rectangle((x, width - pad - bar_height * scale, x + 4 * scale, width - pad), radius=scale, fill=color)
    elif name in {"report", "copy"}:
        draw.rounded_rectangle((pad + 5 * scale, pad, width - pad, width - pad - 5 * scale), radius=2 * scale, outline=color, width=stroke)
        draw.rounded_rectangle((pad, pad + 5 * scale, width - pad - 5 * scale, width - pad), radius=2 * scale, outline=color, width=stroke)
    elif name == "bell":
        draw.arc((pad + 2 * scale, pad, width - pad - 2 * scale, width - pad + 2 * scale), 190, 350, fill=color, width=stroke)
        draw.line((pad + 3 * scale, mid + 5 * scale, width - pad - 3 * scale, mid + 5 * scale), fill=color, width=stroke)
        draw.ellipse((mid - 2 * scale, mid + 6 * scale, mid + 2 * scale, mid + 10 * scale), fill=color)
    elif name == "invalid":
        draw.rounded_rectangle((pad, pad, width - pad, width - pad), radius=2 * scale, outline=color, width=stroke)
        draw.line((pad + 5 * scale, pad + 5 * scale, width - pad - 5 * scale, width - pad - 5 * scale), fill=color, width=stroke)
        draw.line((width - pad - 5 * scale, pad + 5 * scale, pad + 5 * scale, width - pad - 5 * scale), fill=color, width=stroke)
    elif name == "info":
        draw.ellipse((pad, pad, width - pad, width - pad), outline=color, width=stroke)
        draw.ellipse((mid - scale, pad + 4 * scale, mid + scale, pad + 6 * scale), fill=color)
        draw.line((mid, mid - scale, mid, width - pad - 4 * scale), fill=color, width=stroke)
    elif name in {"alert", "clock"}:
        draw.ellipse((pad, pad, width - pad, width - pad), outline=color, width=stroke)
        draw.line((mid, pad + 5 * scale, mid, mid), fill=color, width=stroke)
        draw.line((mid, mid, mid + 5 * scale, mid + 3 * scale), fill=color, width=stroke)
    else:
        draw.rounded_rectangle((pad, pad, width - pad, width - pad), radius=3 * scale, outline=color, width=stroke)

    result = ctk.CTkImage(light_image=image, dark_image=image, size=(size, size))
    ICON_CACHE[key] = result
    return result


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def apply_theme(window):
    window.configure(background=BG)
    style = ttk.Style(window)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", font=(FONT, 10), background=BG, foreground=TEXT)
    style.configure("TFrame", background=BG)
    style.configure("Surface.TFrame", background=SURFACE, relief="flat", borderwidth=0)
    style.configure("Card.TFrame", background="#f7fbff", relief="flat", borderwidth=0)
    style.configure("DangerCard.TFrame", background="#fff1f1", relief="flat", borderwidth=0)
    style.configure("TLabel", background=BG, foreground=TEXT)
    style.configure("Surface.TLabel", background=SURFACE, foreground=TEXT)
    style.configure("Muted.TLabel", background=SURFACE, foreground=MUTED)
    style.configure("CardMuted.TLabel", background="#f7fbff", foreground="#607089")
    style.configure("DangerMuted.TLabel", background="#fff1f1", foreground="#8a4b4b")
    style.configure("Title.TLabel", background=BG, foreground=TEXT, font=(FONT, 17, "bold"))
    style.configure("Section.TLabel", background=SURFACE, foreground=TEXT, font=(FONT, 12, "bold"))
    style.configure("Metric.TLabel", background="#f7fbff", foreground="#0c65c8", font=("Microsoft YaHei UI", 22, "bold"))
    style.configure("DangerMetric.TLabel", background="#fff1f1", foreground="#d84343", font=("Microsoft YaHei UI", 22, "bold"))
    style.configure("TButton", padding=(12, 7), background=SURFACE, bordercolor=BORDER, relief="flat", focusthickness=0)
    style.configure("Primary.TButton", padding=(18, 9), background=TEAL, foreground="#ffffff", font=(FONT, 11, "bold"))
    style.map("Primary.TButton", background=[("active", "#07878a"), ("disabled", "#a7cfd0")])
    style.configure("TEntry", padding=(7, 5), fieldbackground=SURFACE, bordercolor=BORDER, relief="flat")
    style.configure("TCombobox", padding=(7, 5), fieldbackground=SURFACE, bordercolor=BORDER, relief="flat")
    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(18, 8), background="#e7eef7")
    style.map("TNotebook.Tab", background=[("selected", "#ffffff")])
    style.configure("Treeview", rowheight=30, background=SURFACE, fieldbackground=SURFACE, bordercolor=BORDER, relief="flat")
    style.configure("Treeview.Heading", background="#f6f9fc", foreground="#42536c", font=(FONT, 10, "bold"), relief="flat")


class ModernLogTable(ctk.CTkFrame):
    STATUS_COLORS = {
        "信息": ("#e2f1ff", "#1769aa"),
        "成功": ("#dcf7e8", "#147a50"),
        "预警": ("#fff0d4", "#b76400"),
    }

    def __init__(self, parent):
        super().__init__(parent, fg_color=SURFACE, corner_radius=8, border_width=1, border_color="#e2eaf2")
        self._rows = {}
        self._next_id = 0
        header = ctk.CTkFrame(self, fg_color="#f6f9fc", corner_radius=7, height=34)
        header.pack(fill="x", padx=1, pady=(1, 0))
        header.pack_propagate(False)
        header.grid_columnconfigure(2, weight=1)
        for col, text_value, width in ((0, "时间", 100), (1, "状态", 100), (2, "内容", 0)):
            label = ctk.CTkLabel(header, text=text_value, text_color="#52657d", font=(FONT, 10, "bold"), anchor="w", width=width)
            label.grid(row=0, column=col, sticky="ew", padx=(14, 4), pady=5)
        self.body = ctk.CTkScrollableFrame(self, fg_color=SURFACE, corner_radius=0)
        self.body.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        self.body.grid_columnconfigure(2, weight=1)

    def insert(self, _parent, _where, values, tags=()):
        row_id = str(self._next_id)
        self._next_id += 1
        row = ctk.CTkFrame(self.body, fg_color=SURFACE, corner_radius=0, height=32)
        row.grid(row=self._next_id, column=0, columnspan=3, sticky="ew")
        row.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(row, text=str(values[0]), width=100, anchor="w", text_color="#42566e").grid(row=0, column=0, sticky="ew", padx=(14, 4), pady=3)
        status = str(values[1])
        pill_bg, pill_fg = self.STATUS_COLORS.get(status, self.STATUS_COLORS["信息"])
        ctk.CTkLabel(row, text=status, width=58, height=22, corner_radius=11, fg_color=pill_bg, text_color=pill_fg, font=(FONT, 9, "bold")).grid(row=0, column=1, padx=(14, 24), pady=4)
        ctk.CTkLabel(row, text=str(values[2]), anchor="w", justify="left", text_color="#263b53").grid(row=0, column=2, sticky="ew", padx=(0, 12), pady=3)
        ctk.CTkFrame(row, fg_color="#edf2f7", height=1, corner_radius=0).grid(row=1, column=0, columnspan=3, sticky="ew")
        self._rows[row_id] = row
        return row_id

    def get_children(self):
        return tuple(self._rows)

    def delete(self, row_id):
        row = self._rows.pop(str(row_id), None)
        if row:
            row.destroy()

    def yview_moveto(self, fraction):
        self.update_idletasks()
        try:
            self.body._parent_canvas.yview_moveto(fraction)
        except (AttributeError, tk.TclError):
            pass


class ModernComboBox(ctk.CTkComboBox):
    def __init__(self, master, **kwargs):
        width = kwargs.pop("width", None)
        if width is not None:
            kwargs["width"] = max(110, int(width) * 9)
        super().__init__(
            master,
            height=36,
            corner_radius=7,
            border_color=BORDER,
            button_color="#edf4f8",
            button_hover_color="#ddebf1",
            fg_color=SURFACE,
            text_color=TEXT,
            dropdown_fg_color=SURFACE,
            dropdown_text_color=TEXT,
            **kwargs,
        )

    def current(self, index):
        values = self.cget("values")
        if 0 <= index < len(values):
            self.set(values[index])

    def __setitem__(self, key, value):
        if key == "values":
            self.configure(values=value or [""])
            return
        super().__setitem__(key, value)


def ui_button(parent, text, command, primary=False, icon=None, **kwargs):
    width = kwargs.pop("width", 100)
    if width and width < 60:
        width *= 9
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        width=width,
        height=kwargs.pop("height", 36),
        corner_radius=8,
        fg_color=TEAL if primary else SURFACE,
        hover_color="#07878a" if primary else "#eef4f8",
        text_color="#ffffff" if primary else "#194a75",
        border_width=0 if primary else 1,
        border_color=BORDER,
        font=(FONT, 10, "bold"),
        image=ui_icon(icon, "#ffffff" if primary else "#1766a3", 16) if icon else None,
        compound="left",
        **kwargs,
    )


def ui_entry(parent, **kwargs):
    width = kwargs.pop("width", None)
    if width is not None:
        kwargs["width"] = max(110, int(width) * 9)
    return ctk.CTkEntry(
        parent,
        height=36,
        corner_radius=7,
        border_color=BORDER,
        fg_color=SURFACE,
        text_color=TEXT,
        **kwargs,
    )


def ui_check(parent, text, variable, **kwargs):
    return ctk.CTkCheckBox(
        parent,
        text=text,
        variable=variable,
        checkbox_width=20,
        checkbox_height=20,
        corner_radius=5,
        border_color="#aabbd0",
        fg_color=TEAL,
        hover_color="#087f82",
        text_color=TEXT,
        font=(FONT, 10),
        **kwargs,
    )


def ui_label(parent, **kwargs):
    foreground = kwargs.pop("foreground", None)
    style_name = kwargs.pop("style", "")
    if foreground:
        kwargs["text_color"] = foreground
    elif "Muted" in style_name:
        kwargs["text_color"] = MUTED
    else:
        kwargs.setdefault("text_color", TEXT)
    if "Section" in style_name:
        kwargs.setdefault("font", (FONT, 13, "bold"))
    kwargs.setdefault("fg_color", "transparent")
    return ctk.CTkLabel(parent, **kwargs)


def ui_listbox(parent, **kwargs):
    return tk.Listbox(
        parent,
        background=SURFACE,
        foreground="#203047",
        selectbackground="#dff3f4",
        selectforeground="#075f62",
        active_style="none",
        relief="flat",
        borderwidth=0,
        highlightthickness=0,
        font=(FONT, 10),
        **kwargs,
    )


class ReviewTool(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("南方分中心舆情质检辅助工具")
        self.geometry("1280x800")
        self.minsize(1120, 700)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.inspector = tk.StringVar()
        self.status = tk.StringVar(value="")
        self.scheme_var = tk.StringVar()
        self.scheme_options = []
        self.last_report_text = ""
        self.copy_button = None
        self.metric_vars = {}
        self.metric_detail_vars = {}
        self.metrics_frame = None
        self.report_preview = None
        self.drop_title_var = tk.StringVar(value="点击选择文件或拖拽文件到此处")

        self._set_app_icon()
        apply_theme(self)
        self._build_ui()
        self.refresh_schemes()
        self._enable_drop()

    def _set_app_icon(self):
        icon_path = app_dir() / "assets" / "app_icon.ico"
        png_path = app_dir() / "assets" / "app_icon.png"
        if sys.platform.startswith("win") and icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except tk.TclError:
                pass
        elif png_path.exists():
            try:
                self._icon_image = tk.PhotoImage(file=str(png_path))
                self.iconphoto(True, self._icon_image)
            except tk.TclError:
                pass

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="#f2f8fc", corner_radius=0, height=82)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)
        brand = ctk.CTkFrame(header, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="w", padx=22, pady=14)
        icon_path = app_dir() / "assets" / "app_icon.png"
        if icon_path.exists():
            self.header_icon = ctk.CTkImage(Image.open(icon_path), size=(42, 42))
            ctk.CTkLabel(brand, text="", image=self.header_icon, width=46).grid(row=0, column=0, rowspan=2, padx=(0, 10))
        ctk.CTkLabel(brand, text="南方分中心舆情质检辅助工具", text_color=TEXT, font=(FONT, 20, "bold")).grid(row=0, column=1, sticky="w")

        tools = ctk.CTkFrame(header, fg_color="transparent")
        tools.grid(row=0, column=1, sticky="e", padx=22, pady=14)
        tool_items = (
            ("方案配置", "settings", self.open_scheme_config),
            ("专项配置", "special", self.open_special_config),
            ("导入配置", "import", self.import_config),
            ("导出配置", "export", self.export_config),
            ("恢复备份", "restore", self.open_backup_restore),
        )
        for idx, (text_value, icon, command) in enumerate(tool_items):
            self._button(tools, text_value, command, width=112, height=38, icon=icon).grid(row=0, column=idx, padx=(0 if idx == 0 else 7, 0))

        main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        main.grid(row=1, column=0, sticky="nsew", padx=14, pady=(8, 6))
        main.columnconfigure(0, weight=5)
        main.columnconfigure(1, weight=6)
        main.rowconfigure(1, weight=1)

        task = self._panel(main)
        task.grid(row=0, column=0, sticky="nsew", padx=(0, 7), pady=(0, 10))
        task.columnconfigure(0, minsize=86)
        task.columnconfigure(1, weight=1)
        task.columnconfigure(2, minsize=94)
        self._section_title(task, "task", "任务准备").grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(14, 10))
        drop = ctk.CTkFrame(task, fg_color="#f8fbfe", corner_radius=9, border_width=1, border_color="#b9cbe0", height=124, cursor="hand2")
        drop.grid(row=1, column=0, columnspan=3, sticky="ew", padx=16, pady=(0, 12))
        drop.grid_propagate(False)
        drop.columnconfigure(0, weight=1)
        ctk.CTkLabel(drop, text="", image=ui_icon("file", "#7e92aa", 28)).grid(row=0, column=0, pady=(16, 0))
        ctk.CTkLabel(drop, textvariable=self.drop_title_var, text_color="#2d4663", font=(FONT, 11, "bold")).grid(row=1, column=0, pady=(2, 0))
        ctk.CTkLabel(drop, text="支持格式：.xlsx  .xlsm    每次处理一个文件", text_color="#8090a5", font=(FONT, 9)).grid(row=2, column=0, pady=(2, 12))
        for child in (drop, *drop.winfo_children()):
            child.bind("<Button-1>", lambda _event: self.choose_file())

        self._form_row(task, 2, "输出目录", self.output_dir, self.choose_output_dir, "浏览")
        self._field_label(task, "质检方案").grid(row=3, column=0, sticky="e", padx=(16, 10), pady=7)
        self.scheme_combo = ctk.CTkComboBox(task, variable=self.scheme_var, values=[""], state="readonly", height=36, corner_radius=7, border_color=BORDER, button_color="#edf4f8", button_hover_color="#ddebf1", text_color=TEXT, dropdown_fg_color=SURFACE)
        self.scheme_combo.grid(row=3, column=1, sticky="ew", pady=7)
        self._button(task, "方案管理", self.open_scheme_config, width=94, icon="settings").grid(row=3, column=2, sticky="ew", padx=(10, 16), pady=7)
        self._field_label(task, "质检人员").grid(row=4, column=0, sticky="e", padx=(16, 10), pady=7)
        inspector_row = ctk.CTkFrame(task, fg_color="transparent")
        inspector_row.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(0, 16), pady=7)
        inspector_row.columnconfigure(0, weight=1)
        self.inspector_combo = ctk.CTkComboBox(inspector_row, variable=self.inspector, values=[""], height=36, corner_radius=7, border_color=BORDER, button_color="#edf4f8", button_hover_color="#ddebf1", text_color=TEXT, dropdown_fg_color=SURFACE)
        self.inspector_combo.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(inspector_row, text="用于结果记录与报送内容标注", text_color="#8795a8", font=(FONT, 9)).grid(row=0, column=1, padx=(10, 0))
        action_row = ctk.CTkFrame(task, fg_color="transparent")
        action_row.grid(row=5, column=0, columnspan=3, sticky="ew", padx=16, pady=(12, 16))
        action_row.columnconfigure(0, weight=7)
        action_row.columnconfigure(1, weight=3)
        self.run_button = self._button(action_row, "开始质检", self.run, primary=True, height=44, font=(FONT, 12, "bold"), icon="play")
        self.run_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._button(action_row, "清空文件", self.clear_file, height=44, icon="trash").grid(row=0, column=1, sticky="ew", padx=(6, 0))

        result = self._panel(main)
        result.grid(row=0, column=1, sticky="nsew", padx=(7, 0), pady=(0, 10))
        result.columnconfigure(0, weight=1)
        top_result = ctk.CTkFrame(result, fg_color="transparent")
        top_result.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 10))
        top_result.columnconfigure(0, weight=1)
        self._section_title(top_result, "chart", "本次质检结果").grid(row=0, column=0, sticky="w")
        self.task_time_var = tk.StringVar(value="任务时间：-")
        ctk.CTkLabel(top_result, textvariable=self.task_time_var, text_color="#8795a8", font=(FONT, 9)).grid(row=0, column=1, sticky="e")
        self.metrics_frame = ctk.CTkFrame(result, fg_color="transparent")
        self.metrics_frame.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))
        self._render_metric_cards([
            {"title": "月度专项", "key": "special", "value": "0"},
            {"title": "舆情提醒", "key": "reminder", "value": "0"},
            {"title": "无效舆情", "key": "invalid", "value": "0"},
            {"title": "超时预警", "key": "overtime", "value": "0", "danger": True},
        ])
        report_head = ctk.CTkFrame(result, fg_color="transparent")
        report_head.grid(row=2, column=0, sticky="ew", padx=16)
        report_head.columnconfigure(0, weight=1)
        self._section_title(report_head, "report", "报送内容").grid(row=0, column=0, sticky="w")
        self.copy_button = self._button(
            report_head,
            "复制报送内容",
            self.copy_report,
            primary=True,
            width=148,
            state="disabled",
            icon="copy",
        )
        self.copy_button.grid(row=0, column=1, rowspan=2, sticky="e")
        ctk.CTkLabel(report_head, text="已根据本次质检结果生成报送内容，可一键复制用于报送。", text_color="#7b8ba0", font=(FONT, 9)).grid(row=1, column=0, sticky="w", pady=(3, 0))
        self.report_preview = ctk.CTkTextbox(result, height=72, corner_radius=8, border_width=1, border_color="#d4dfeb", fg_color="#fbfdff", text_color="#263b53", wrap="word", font=(FONT, 10))
        self.report_preview.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 16))
        self.report_preview.insert("1.0", "处理完成后将在这里显示日报送内容。")
        self.report_preview.configure(state="disabled")

        logs = ctk.CTkTabview(main, fg_color=SURFACE, segmented_button_fg_color="#edf3f8", segmented_button_selected_color=SURFACE, segmented_button_selected_hover_color=SURFACE, segmented_button_unselected_color="#edf3f8", segmented_button_unselected_hover_color="#e1ebf3", text_color=TEXT, corner_radius=10, border_width=1, border_color="#d8e3ed")
        logs.grid(row=1, column=0, columnspan=2, sticky="nsew")
        logs.add("处理日志")
        logs.add("质检报送")
        logs.add("校验提醒")
        for tab_name, icon_name in (("处理日志", "log"), ("质检报送", "report"), ("校验提醒", "alert")):
            try:
                logs._segmented_button._buttons_dict[tab_name].configure(
                    image=ui_icon(icon_name, "#1766a3", 15), compound="left"
                )
            except (AttributeError, KeyError, tk.TclError):
                pass
        log_tab = logs.tab("处理日志")
        report_tab = logs.tab("质检报送")
        alert_tab = logs.tab("校验提醒")
        log_tab.columnconfigure(0, weight=1)
        log_tab.rowconfigure(1, weight=1)
        log_tools = ctk.CTkFrame(log_tab, fg_color="transparent")
        log_tools.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        log_tools.columnconfigure(0, weight=1)
        self._button(log_tools, "清空日志", self.clear_log, width=92, height=30).grid(row=0, column=1, sticky="e")
        self.log = self._make_log_table(log_tab, row=1)
        report_tab.columnconfigure(0, weight=1)
        report_tab.rowconfigure(0, weight=1)
        self.report_detail = ctk.CTkTextbox(report_tab, corner_radius=8, border_width=1, border_color="#e1e9f1", fg_color=SURFACE, text_color="#263b53", wrap="word", font=(FONT, 10))
        self.report_detail.grid(row=0, column=0, sticky="nsew")
        self.report_detail.configure(state="disabled")
        alert_tab.columnconfigure(0, weight=1)
        alert_tab.rowconfigure(1, weight=1)
        alert_tools = ctk.CTkFrame(alert_tab, fg_color="transparent")
        alert_tools.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        alert_tools.columnconfigure(0, weight=1)
        self._button(alert_tools, "清空日志", self.clear_log, width=92, height=30).grid(row=0, column=1, sticky="e")
        self.alert_log = self._make_log_table(alert_tab, row=1)

        footer = ctk.CTkFrame(self, fg_color="#f2f8fc", corner_radius=0, height=34)
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(1, weight=1)
        self.status.set("就绪")
        ctk.CTkLabel(footer, textvariable=self.status, text_color="#42566e", font=(FONT, 9)).grid(row=0, column=0, sticky="w", padx=18)
        self.footer_scheme_var = tk.StringVar(value="当前方案：-")
        self.footer_time_var = tk.StringVar(value="时间：" + datetime.now().strftime("%Y-%m-%d %H:%M"))
        ctk.CTkLabel(footer, textvariable=self.footer_scheme_var, text_color="#596b82", font=(FONT, 9)).grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(footer, text="  |  ", text_color="#a4afbd").grid(row=0, column=2)
        ctk.CTkLabel(footer, textvariable=self.footer_time_var, text_color="#596b82", font=(FONT, 9)).grid(row=0, column=3, sticky="e", padx=(0, 18))
        self._log("请选择源文件后开始质检。")

    def _panel(self, parent):
        return ctk.CTkFrame(parent, fg_color=SURFACE, corner_radius=10, border_width=1, border_color="#dce6ef")

    @staticmethod
    def _section_title(parent, icon, text):
        return ctk.CTkLabel(
            parent,
            text=text,
            image=ui_icon(icon, "#1766a3", 19),
            compound="left",
            text_color="#14345f",
            font=(FONT, 13, "bold"),
        )

    @staticmethod
    def _field_label(parent, text):
        return ctk.CTkLabel(parent, text=text, text_color="#304057", width=72, anchor="e", font=(FONT, 10, "bold"))

    @staticmethod
    def _button(parent, text, command, primary=False, width=100, height=34, font=None, state="normal", icon=None):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=height,
            corner_radius=8,
            fg_color=TEAL if primary else SURFACE,
            hover_color="#07888b" if primary else "#f1f6fa",
            text_color="#ffffff" if primary else "#194a75",
            border_width=0 if primary else 1,
            border_color="#c9d7e6",
            font=font or (FONT, 10, "bold"),
            image=ui_icon(icon, "#ffffff" if primary else "#1766a3", 16) if icon else None,
            compound="left",
            state=state,
        )

    def _form_row(self, parent, row, label, var, command, button_text):
        self._field_label(parent, label).grid(row=row, column=0, sticky="e", padx=(16, 10), pady=7)
        ctk.CTkEntry(parent, textvariable=var, height=36, corner_radius=7, border_color=BORDER, fg_color=SURFACE, text_color=TEXT).grid(row=row, column=1, sticky="ew", pady=7)
        self._button(parent, button_text, command, width=94).grid(row=row, column=2, sticky="ew", padx=(10, 16), pady=7)

    def _render_metric_cards(self, cards):
        for widget in self.metrics_frame.winfo_children():
            widget.destroy()
        self.metric_vars = {}
        self.metric_detail_vars = {}
        count = max(1, len(cards))
        for idx in range(count):
            self.metrics_frame.columnconfigure(idx, weight=1, uniform="metric")
        for idx, card in enumerate(cards):
            self._metric_card(
                self.metrics_frame,
                idx,
                card.get("title", "质检计划"),
                card.get("key", f"metric_{idx}"),
                card.get("value", "0"),
                card.get("danger") is True,
                card.get("involved"),
                card.get("attention"),
            )

    def _metric_card(self, parent, col, title, key, value, danger=False, involved=None, attention=None):
        colors = [
            ("#eaf5ff", "#1474d4", "monthly"),
            ("#e9faf6", "#008b8e", "bell"),
            ("#f2f5f8", "#425168", "invalid"),
            ("#fff0f1", "#e0444e", "clock"),
        ]
        bg, fg, icon = colors[3 if danger else col % 3]
        frame = ctk.CTkFrame(parent, fg_color=bg, corner_radius=9, border_width=0)
        frame.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 6, 0))
        frame.grid_columnconfigure(0, weight=1)
        self.metric_vars[key] = tk.StringVar(value=str(value))
        detail_var = tk.StringVar(value=f"涉及条数：{value if involved is None else involved}\n需关注：{0 if attention is None else attention}")
        self.metric_detail_vars[key] = detail_var
        card_head = ctk.CTkFrame(frame, fg_color="transparent")
        card_head.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        card_head.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(card_head, text="", image=ui_icon(icon, fg, 20), width=24).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkLabel(card_head, text=title, text_color="#304057", font=(FONT, 10, "bold"), anchor="w").grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(card_head, text="", image=ui_icon("info", "#8b9bb0", 13), width=16).grid(row=0, column=2)
        ctk.CTkLabel(frame, textvariable=self.metric_vars[key], text_color=fg, font=(FONT, 25, "bold")).grid(row=1, column=0, sticky="ew", padx=12, pady=(5, 1))
        ctk.CTkLabel(frame, textvariable=detail_var, text_color="#66758a", font=(FONT, 9), justify="center").grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))

    def _make_log_table(self, parent, row=0):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(row, weight=1)
        table = ModernLogTable(parent)
        table.grid(row=row, column=0, sticky="nsew")
        return table

    def _enable_drop(self):
        if not DND_FILES:
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
        self.drop_title_var.set(path.name)
        self.status.set("已选择源文件")

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择舆情质检明细",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)
            self.drop_title_var.set(Path(path).name)
            self.status.set("已选择源文件")

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)
            self.status.set("已选择输出目录")

    def clear_file(self):
        self.input_path.set("")
        self.drop_title_var.set("点击选择文件或拖拽文件到此处")
        self.status.set("已清空源文件")

    def refresh_schemes(self):
        config = load_config()
        schemes = config.get("schemes") or {DEFAULT_SCHEME_ID: DEFAULT_SCHEME}
        active_id = config.get("active_scheme_id") or DEFAULT_SCHEME_ID
        self.scheme_options = [(scheme_id, scheme.get("name") or scheme_id) for scheme_id, scheme in schemes.items()]
        names = [name for _scheme_id, name in self.scheme_options]
        self.scheme_combo.configure(values=names or [""])
        selected_index = 0
        for idx, (scheme_id, _name) in enumerate(self.scheme_options):
            if scheme_id == active_id:
                selected_index = idx
                break
        if self.scheme_options:
            self.scheme_combo.set(self.scheme_options[selected_index][1])
        if hasattr(self, "footer_scheme_var"):
            self.footer_scheme_var.set("当前方案：" + (self.scheme_var.get() or "-"))

    def selected_scheme_id(self):
        selected_name = self.scheme_var.get()
        for scheme_id, name in self.scheme_options:
            if name == selected_name:
                return scheme_id
        return DEFAULT_SCHEME_ID

    def open_scheme_config(self):
        SchemeConfigWindow(self)

    def open_special_config(self):
        SpecialConfigWindow(self)

    def export_config(self):
        config = load_config()
        scheme_name = self.scheme_var.get() or "当前方案"
        default_name = "舆情质检配置_" + datetime.now().strftime("%Y%m%d") + "_" + scheme_name + ".json"
        path = filedialog.asksaveasfilename(
            title="导出配置",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("配置文件", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        self.status.set("配置已导出")
        messagebox.showinfo("完成", "配置已导出。")

    def import_config(self):
        path = filedialog.askopenfilename(
            title="导入配置",
            filetypes=[("配置文件", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                imported = json.load(f)
            if not isinstance(imported, dict):
                raise ValueError("配置文件格式不正确")
            if "schemes" not in imported and "monthly_special_plans" not in imported:
                raise ValueError("未识别到有效的质检配置内容")
            save_config(imported)
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            return
        self.refresh_schemes()
        self.status.set("配置已导入")
        messagebox.showinfo("完成", "配置已导入，已自动备份导入前的旧配置。")

    def open_backup_restore(self):
        BackupRestoreWindow(self)

    def _validate_before_run(self, input_path, scheme_id):
        result = validate_config_for_file(input_path, scheme_id)
        if not result["ok"]:
            messagebox.showerror("配置校验未通过", "\n".join(result["errors"][:12]))
            self._log("配置校验未通过：" + "；".join(result["errors"]), "red")
            return False
        summary = f"配置校验通过：当前方案【{result['scheme_name']}】，识别源表 {result['column_count']} 列。"
        self._log(summary)
        if result["warnings"]:
            warning_text = "\n".join(result["warnings"][:12])
            return messagebox.askyesno("配置校验提醒", warning_text + "\n\n是否继续处理？")
        return True

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
        scheme_id = self.selected_scheme_id()
        if not self._validate_before_run(input_path, scheme_id):
            return

        self.last_report_text = ""
        self.copy_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.status.set("正在处理，请稍候...")
        self.task_time_var.set("任务时间：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self._set_report_text("正在处理，请稍候...")
        self._set_detail_text(self.report_detail, "")
        self._render_metric_cards([
            {"title": "月度专项", "key": "special", "value": "0"},
            {"title": "舆情提醒", "key": "reminder", "value": "0"},
            {"title": "无效舆情", "key": "invalid", "value": "0"},
            {"title": "超时预警", "key": "overtime", "value": "0", "danger": True},
        ])
        self._log("开始处理：" + input_path)
        inspector = self.inspector.get()
        threading.Thread(target=self._run_worker, args=(input_path, output_dir, scheme_id, inspector), daemon=True).start()

    def _run_worker(self, input_path, output_dir, scheme_id, inspector):
        try:
            result = process_file(
                input_path,
                output_dir,
                inspector,
                progress_callback=self._thread_log,
                scheme_id=scheme_id,
                save_output=False,
            )
            self.after(0, self._run_success, result)
        except Exception as exc:
            self.after(0, self._run_failed, exc)

    def _thread_log(self, text):
        self.after(0, self._log, text)

    def _run_success(self, result):
        self.run_button.configure(state="normal")
        self.status.set("处理完成")
        self.last_report_text = result["report_text"]
        self.copy_button.configure(state="normal")
        cards = []
        for idx, item in enumerate(result.get("plan_results") or []):
            plan = item.get("plan") or {}
            if not plan.get("output_sheet", True):
                continue
            title = plan.get("name") or "质检计划"
            if plan.get("role") == "monthly_special":
                title = "月度专项"
            elif plan.get("role") == "reminder_review":
                title = "舆情提醒"
            elif plan.get("role") == "invalid_review":
                title = "无效舆情"
            row_count = len(item.get("rows") or [])
            cards.append({
                "title": title,
                "key": f"plan_{idx}",
                "value": row_count,
                "involved": row_count,
                "attention": len(item.get("overtime_ids") or []),
            })
        cards.append({
            "title": "超时预警",
            "key": "overtime",
            "value": result["overtime_count"],
            "involved": result["overtime_count"],
            "attention": result["overtime_count"],
            "danger": True,
        })
        self._render_metric_cards(cards)
        self._log("专项质检：" + result["special_name"])
        self._log(f"专项命中：{result['special_count']} 条")
        self._log(f"舆情提醒复核：{result['reminder_count']} 条")
        self._log(f"无效复核抽样：{result['invalid_count']} 条")
        if result["overtime_count"]:
            text = "超时预警：" + str(result["overtime_count"]) + " 条，编号：" + "、".join(result["overtime_ids"])
            self._log(text, "red")
            messagebox.showwarning("超时预警", text)
        else:
            self._log("超时预警：未查询到超时舆情")
        self._log("统计日报文本：" + result["report_text"])
        self._log("质检报送内容：" + result["submission_text"])
        self._set_report_text(result["report_text"])
        self._set_detail_text(self.report_detail, result["submission_text"] + "\n\n" + result["report_text"])
        if messagebox.askyesno("导出质检明细", "是否导出质检明细文件？"):
            try:
                output_path = export_quality_file(result, progress_callback=self._log)
            except Exception as exc:
                self.status.set("导出失败")
                self._log("导出失败：" + str(exc), "red")
                messagebox.showerror("导出失败", str(exc))
                return
            self._log("输出文件：" + output_path)
            messagebox.showinfo("完成", "舆情质检明细已生成。")
        else:
            self._log("用户选择不导出质检明细。")
            messagebox.showinfo("完成", "处理完成，未导出质检明细。")

    def _run_failed(self, exc):
        self.run_button.configure(state="normal")
        self.status.set("处理失败")
        self._log("处理失败：" + str(exc), "red")
        messagebox.showerror("处理失败", str(exc))

    def copy_report(self):
        if not self.last_report_text:
            messagebox.showinfo("提示", "暂无可复制的日报送内容")
            return
        self.clipboard_clear()
        self.clipboard_append(self.last_report_text)
        self.status.set("日报送内容已复制")

    def clear_log(self):
        for tree in (self.log, self.alert_log):
            for item in tree.get_children():
                tree.delete(item)

    def _set_report_text(self, text):
        if not self.report_preview:
            return
        self._set_detail_text(self.report_preview, text)

    @staticmethod
    def _set_detail_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def _log(self, text, tag=None):
        status = "预警" if tag == "red" or "超时预警" in text or "失败" in text else "信息"
        if "完成" in text or "通过" in text or "成功" in text:
            status = "成功"
        values = (datetime.now().strftime("%H:%M:%S"), status, text)
        row_tag = "red" if status == "预警" else ("success" if status == "成功" else "")
        self.log.insert("", "end", values=values, tags=(row_tag,))
        self.log.yview_moveto(1)
        if status == "预警":
            self.alert_log.insert("", "end", values=values, tags=("red",))
            self.alert_log.yview_moveto(1)


class BackupRestoreWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("恢复配置备份")
        self.geometry("620x420")
        self.minsize(560, 360)
        self.transient(parent)
        apply_theme(self)
        self.backup_dir = app_dir() / "config_backups"
        self.backups = []
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        card = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=10, border_width=1, border_color=BORDER)
        card.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ui_label(card, text="配置备份", style="Section.TLabel").grid(row=0, column=0, sticky="w", padx=16, pady=(14, 0))
        ui_label(card, text="选择一个历史备份，恢复前会自动保存当前配置。", foreground="#546179").grid(
            row=1, column=0, sticky="w", padx=16, pady=(4, 8)
        )
        list_shell = ctk.CTkFrame(card, fg_color=SURFACE, corner_radius=9, border_width=1, border_color=BORDER)
        list_shell.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 10))
        list_shell.columnconfigure(0, weight=1)
        list_shell.rowconfigure(0, weight=1)
        self.listbox = ui_listbox(list_shell, exportselection=False)
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 16))
        ui_button(actions, text="刷新", command=self.refresh).pack(side="left")
        ui_button(actions, text="恢复选中备份", command=self.restore_selected, primary=True, width=140).pack(side="right")

    def refresh(self):
        self.listbox.delete(0, "end")
        if not self.backup_dir.exists():
            return
        self.backups = sorted(self.backup_dir.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
        for path in self.backups:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            self.listbox.insert("end", f"{modified}  {path.name}")

    def restore_selected(self):
        if not self.listbox.curselection():
            messagebox.showinfo("提示", "请先选择一个备份。")
            return
        path = self.backups[self.listbox.curselection()[0]]
        if not messagebox.askyesno("确认恢复", "确认恢复选中的配置备份？"):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                restored = json.load(f)
            if not isinstance(restored, dict):
                raise ValueError("备份文件格式不正确")
            save_config(restored)
        except Exception as exc:
            messagebox.showerror("恢复失败", str(exc))
            return
        self.parent.refresh_schemes()
        self.parent.status.set("配置备份已恢复")
        messagebox.showinfo("完成", "配置备份已恢复。")
        self.destroy()


class SchemeConfigWindow(tk.Toplevel):
    FIELD_OPTIONS = [
        ("keep", "基础保留"),
        ("company", "公司筛选"),
        ("send_time", "派发时间"),
        ("process_time", "处理时间"),
        ("event_nature", "事件性质"),
        ("valid_scope", "舆情范围"),
        ("marketing", "营销类"),
        ("reminder", "舆情提醒"),
        ("duplicate", "重复事件"),
        ("event_category", "事件分类"),
        ("id", "编号"),
        ("special_targets", "专项关键词匹配"),
    ]

    VALUE_OPTIONS = [
        ("valid_yes", "符合舆情范围：是", "valid_scope", "是"),
        ("valid_no", "符合舆情范围：否", "valid_scope", "否"),
        ("marketing_yes", "营销类：是", "marketing", "是"),
        ("marketing_no", "营销类：否", "marketing", "否"),
        ("duplicate_yes", "重复事件：是", "duplicate", "是"),
        ("duplicate_no", "重复事件：否", "duplicate", "否"),
        ("negative_event", "负面事件值", "event_nature", "负面事件"),
        ("positive_event", "正面事件值", "event_nature", "正面事件"),
        ("reminder", "舆情提醒值", "reminder", "舆情提醒"),
        ("livelihood_event", "民生事件值", "event_category", "民生类舆情"),
    ]

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("方案配置")
        self.geometry("1220x820")
        self.minsize(1080, 720)
        self.transient(parent)
        apply_theme(self)

        self.config_data = load_config()
        self.schemes = copy.deepcopy(self.config_data.get("schemes") or {DEFAULT_SCHEME_ID: DEFAULT_SCHEME})
        self.current_scheme_id = None
        self.header_choices = []
        self.name_var = tk.StringVar()
        self.config_version_var = tk.StringVar(value=self.config_data.get("config_version", ""))
        self.config_remark_var = tk.StringVar(value=self.config_data.get("config_remark", ""))
        self.sample_rate_var = tk.StringVar()
        self.sample_min_var = tk.StringVar()
        self.overtime_var = tk.StringVar()
        self.field_row_vars = []
        self.value_row_vars = []
        self.plan_row_vars = []
        self.field_rows_frame = None
        self.value_rows_frame = None
        self.plan_rows_frame = None

        self._build_ui()
        self.refresh_scheme_list()

    def _build_ui(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=10, border_width=1, border_color=BORDER)
        left.grid(row=0, column=0, sticky="ns", padx=(14, 7), pady=14)
        ctk.CTkLabel(left, text="质检方案", text_color=TEXT, font=(FONT, 13, "bold")).pack(anchor="w", padx=14, pady=(14, 0))
        scheme_shell = ctk.CTkFrame(left, fg_color=SURFACE, corner_radius=8, border_width=1, border_color=BORDER)
        scheme_shell.pack(fill="both", expand=True, padx=14, pady=(8, 10))
        self.scheme_list = ui_listbox(scheme_shell, height=14, exportselection=False, width=20)
        self.scheme_list.pack(fill="both", expand=True, padx=8, pady=8)
        self.scheme_list.bind("<<ListboxSelect>>", self.on_scheme_select)
        ui_button(left, text="新增方案", command=self.add_scheme).pack(fill="x", padx=14, pady=(0, 6))
        ui_button(left, text="复制方案", command=self.copy_scheme).pack(fill="x", padx=14, pady=(0, 6))
        ui_button(left, text="删除方案", command=self.delete_scheme).pack(fill="x", padx=14, pady=(0, 14))

        right = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=10, border_width=1, border_color=BORDER)
        right.grid(row=0, column=1, sticky="nsew", padx=(7, 14), pady=14)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        top = ctk.CTkFrame(right, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 10))
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)
        ui_label(top, text="方案名称").grid(row=0, column=0, sticky="e", padx=(0, 8))
        ui_entry(top, textvariable=self.name_var).grid(row=0, column=1, sticky="ew")
        ui_label(top, text="配置版本").grid(row=0, column=2, sticky="e", padx=(10, 8))
        ui_entry(top, textvariable=self.config_version_var).grid(row=0, column=3, sticky="ew")
        ui_label(top, text="配置备注").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=(8, 0))
        ui_entry(top, textvariable=self.config_remark_var).grid(row=1, column=1, columnspan=3, sticky="ew", pady=(8, 0))
        ui_button(top, text="读取源文件表头", command=self.load_headers, width=140).grid(row=0, column=4, rowspan=2, padx=(10, 0))

        notebook = ctk.CTkTabview(
            right,
            fg_color="#f8fbfd",
            corner_radius=9,
            border_width=1,
            border_color="#dce6ef",
            segmented_button_selected_color=TEAL,
            segmented_button_selected_hover_color=TEAL_HOVER,
        )
        notebook.grid(row=1, column=0, sticky="nsew", padx=16)
        notebook.add("字段映射")
        notebook.add("质检计划")
        notebook.add("公司名单")
        fields_tab = notebook.tab("字段映射")
        values_tab = notebook.tab("质检计划")
        company_tab = notebook.tab("公司名单")

        fields_tab.columnconfigure(0, weight=1)
        fields_tab.rowconfigure(1, weight=1)
        field_actions = ctk.CTkFrame(fields_tab, fg_color="transparent")
        field_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ui_button(field_actions, text="新增字段", command=self.add_field_row).pack(side="left")
        ui_label(
            field_actions,
            text="用途决定程序逻辑，显示名称和列可按源表调整；专项关键词匹配列可填多个，如 C,K。",
            foreground="#546179",
        ).pack(side="left", padx=(10, 0))
        self.field_rows_frame = self._make_scroll_area(fields_tab, row=1)

        values_tab.columnconfigure(0, weight=1)
        values_tab.rowconfigure(1, weight=1)
        value_actions = ctk.CTkFrame(values_tab, fg_color="transparent")
        value_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ui_button(value_actions, text="新增质检计划", command=self.add_plan_row, width=130).pack(side="left")
        ui_label(
            value_actions,
            text="每个计划可单独配置筛选方式、导出、抽样和超时检查。",
            foreground="#546179",
        ).pack(side="left", padx=(10, 0))
        self.plan_rows_frame = self._make_scroll_area(values_tab, row=1)

        company_tab.columnconfigure(0, weight=1)
        company_tab.rowconfigure(0, weight=1)
        self.company_text = ctk.CTkTextbox(company_tab, height=360, wrap="word", corner_radius=8, border_width=1, border_color=BORDER, fg_color=SURFACE, text_color=TEXT)
        self.company_text.grid(row=0, column=0, sticky="nsew")
        ui_label(company_tab, text="每行一个公司名称；公司筛选列的值在名单中才会保留。", foreground="#546179").grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )

        actions = ctk.CTkFrame(right, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=16, pady=(12, 16))
        ui_button(actions, text="保存当前方案", command=self.save_current_scheme, width=130).pack(side="left")
        ui_button(actions, text="设为默认方案", command=self.set_active_scheme, width=130).pack(side="left", padx=8)
        ui_button(actions, text="恢复当前方案默认值", command=self.reset_current_scheme, width=170).pack(side="left")
        ui_button(actions, text="保存全部并关闭", command=self.save_all_and_close, primary=True, width=150).pack(side="right")

    def refresh_scheme_list(self):
        self.scheme_list.delete(0, "end")
        ids = list(self.schemes.keys())
        for scheme_id in ids:
            self.scheme_list.insert("end", self.schemes[scheme_id].get("name") or scheme_id)
        active_id = self.config_data.get("active_scheme_id") or DEFAULT_SCHEME_ID
        selected_index = ids.index(active_id) if active_id in ids else 0
        self.scheme_list.selection_set(selected_index)
        self.load_scheme(ids[selected_index])

    def on_scheme_select(self, _event=None):
        if not self.scheme_list.curselection():
            return
        self.save_current_scheme(show_message=False)
        scheme_id = list(self.schemes.keys())[self.scheme_list.curselection()[0]]
        self.load_scheme(scheme_id)

    def load_scheme(self, scheme_id):
        self.current_scheme_id = scheme_id
        scheme = self.schemes[scheme_id]
        self.name_var.set(scheme.get("name") or scheme_id)
        self.field_row_vars = []
        self.value_row_vars = []
        self.plan_row_vars = []
        for widget in self.field_rows_frame.winfo_children():
            widget.destroy()
        for widget in self.plan_rows_frame.winfo_children():
            widget.destroy()
        for item in self._scheme_field_items(scheme):
            self.add_field_row(item)
        for item in self._scheme_review_plans(scheme):
            self.add_plan_row(item)
        self.sample_rate_var.set(str(scheme.get("invalid_sample_rate", 0.2)))
        self.sample_min_var.set(str(scheme.get("invalid_sample_min", 1)))
        self.overtime_var.set(str(scheme.get("overtime_threshold_minutes", 20)))
        self.company_text.delete("1.0", "end")
        values = scheme.get("values") or DEFAULT_SCHEME["values"]
        self.company_text.insert("1.0", "\n".join(values.get("company_names") or sorted(DEFAULT_SCHEME["values"]["company_names"])))

    def save_current_scheme(self, show_message=True):
        if not self.current_scheme_id:
            return True
        field_items = self._collect_field_items()
        review_plans = self._collect_review_plans()
        if review_plans is None:
            return False
        value_rules = self._value_rules_from_plans(review_plans)
        fields = {}
        for item in field_items:
            if not item["enabled"]:
                continue
            if item["key"] == "special_targets":
                fields[item["key"]] = [part.strip() for part in item["column"].replace("，", ",").replace("、", ",").split(",") if part.strip()]
            else:
                fields[item["key"]] = item["column"]
        values = {item["key"]: item["value"] for item in value_rules if item["enabled"]}
        values["company_names"] = [line.strip() for line in self.company_text.get("1.0", "end").splitlines() if line.strip()]
        try:
            sample_rate = float(self.sample_rate_var.get() or 0.2)
            sample_min = int(self.sample_min_var.get() or 1)
            overtime = float(self.overtime_var.get() or 20)
        except ValueError:
            messagebox.showerror("配置错误", "抽样比例、最少条数和超时阈值必须是数字。")
            return False
        new_scheme = {
            "id": self.current_scheme_id,
            "name": self.name_var.get().strip() or self.current_scheme_id,
            "fields": fields,
            "field_items": field_items,
            "values": values,
            "value_rules": value_rules,
            "review_plans": review_plans,
            "invalid_sample_rate": sample_rate,
            "invalid_sample_min": sample_min,
            "overtime_threshold_minutes": overtime,
        }
        errors = self._validate_scheme_for_save(new_scheme)
        if errors:
            messagebox.showerror("配置错误", "\n".join(errors[:12]))
            return False
        self.schemes[self.current_scheme_id] = new_scheme
        if show_message:
            messagebox.showinfo("已保存", "当前方案已暂存")
        return True

    def add_scheme(self):
        self.save_current_scheme(show_message=False)
        base = "scheme"
        idx = 1
        while f"{base}_{idx}" in self.schemes:
            idx += 1
        scheme_id = f"{base}_{idx}"
        scheme = copy_json(DEFAULT_SCHEME)
        scheme["id"] = scheme_id
        scheme["name"] = f"新方案{idx}"
        self.schemes[scheme_id] = scheme
        self.refresh_scheme_list()
        keys = list(self.schemes.keys())
        self.scheme_list.selection_clear(0, "end")
        self.scheme_list.selection_set(keys.index(scheme_id))
        self.load_scheme(scheme_id)

    def copy_scheme(self):
        if not self.current_scheme_id:
            return
        self.save_current_scheme(show_message=False)
        base_id = self.current_scheme_id
        idx = 1
        while f"{base_id}_copy_{idx}" in self.schemes:
            idx += 1
        scheme_id = f"{base_id}_copy_{idx}"
        scheme = copy_json(self.schemes[base_id])
        scheme["id"] = scheme_id
        scheme["name"] = scheme.get("name", base_id) + " 副本"
        self.schemes[scheme_id] = scheme
        self.refresh_scheme_list()

    def delete_scheme(self):
        if self.current_scheme_id == DEFAULT_SCHEME_ID:
            messagebox.showwarning("提示", "默认结构方案不能删除")
            return
        if not messagebox.askyesno("确认删除", "确认删除当前方案？"):
            return
        del self.schemes[self.current_scheme_id]
        if self.config_data.get("active_scheme_id") == self.current_scheme_id:
            self.config_data["active_scheme_id"] = DEFAULT_SCHEME_ID
        self.refresh_scheme_list()

    def reset_current_scheme(self):
        if not self.current_scheme_id:
            return
        default = copy_json(DEFAULT_SCHEME)
        default["id"] = self.current_scheme_id
        if self.current_scheme_id != DEFAULT_SCHEME_ID:
            default["name"] = self.name_var.get().strip() or self.current_scheme_id
        self.schemes[self.current_scheme_id] = default
        self.load_scheme(self.current_scheme_id)
        messagebox.showinfo("已恢复", "当前方案已恢复为默认结构")

    def set_active_scheme(self):
        self.save_current_scheme(show_message=False)
        self.config_data["active_scheme_id"] = self.current_scheme_id
        messagebox.showinfo("已设置", "当前方案已设为默认运行方案")

    def save_all_and_close(self):
        if not self.save_current_scheme(show_message=False):
            return
        if not self.config_version_var.get().strip():
            self.config_version_var.set(datetime.now().strftime("%Y-%m-%d-001"))
        self.config_data["config_version"] = self.config_version_var.get().strip()
        self.config_data["config_remark"] = self.config_remark_var.get().strip()
        self.config_data["schemes"] = self.schemes
        if not self.config_data.get("active_scheme_id"):
            self.config_data["active_scheme_id"] = self.current_scheme_id or DEFAULT_SCHEME_ID
        save_config(self.config_data)
        self.parent.refresh_schemes()
        self.parent.status.set("质检方案配置已保存")
        self.destroy()

    def load_headers(self):
        path = self.parent.input_path.get().strip()
        if not path or not Path(path).exists():
            path = filedialog.askopenfilename(
                title="选择用于读取表头的 Excel",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
            )
        if not path:
            return
        try:
            wb = load_config_headers(path)
            self.header_choices = wb
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        self._refresh_column_choices()
        messagebox.showinfo("完成", "已读取表头，可在字段映射中选择。")

    def _make_scroll_area(self, parent, row):
        frame = ctk.CTkScrollableFrame(
            parent,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color="#bac8d8",
            scrollbar_button_hover_color="#98abbe",
        )
        frame.columnconfigure(0, weight=1)
        frame.grid(row=row, column=0, sticky="nsew")
        return frame

    def _field_display(self, key):
        labels = dict(self.FIELD_OPTIONS)
        return labels.get(key, "自定义字段")

    def _value_display(self, key):
        labels = {item[0]: item[1] for item in self.VALUE_OPTIONS}
        return labels.get(key, "自定义规则")

    def _parse_field_key(self, display):
        reverse = {label: key for key, label in self.FIELD_OPTIONS}
        return reverse.get(str(display or "").strip(), "custom_field")

    def _parse_value_key(self, display):
        reverse = {label: key for key, label, _field_key, _default in self.VALUE_OPTIONS}
        return reverse.get(str(display or "").strip(), "custom_rule")

    def _column_choices(self):
        return self.header_choices or []

    def _scheme_field_items(self, scheme):
        items = copy_json(scheme.get("field_items") or [])
        if not items:
            fields = scheme.get("fields") or DEFAULT_SCHEME["fields"]
            labels = {item["key"]: item["label"] for item in DEFAULT_SCHEME["field_items"]}
            for key, column in fields.items():
                if isinstance(column, list):
                    column = ",".join(str(item) for item in column)
                items.append({"key": key, "label": labels.get(key, key), "column": str(column), "enabled": True})
        return items or copy_json(DEFAULT_SCHEME["field_items"])

    def _scheme_value_rules(self, scheme):
        items = copy_json(scheme.get("value_rules") or [])
        if not items:
            values = scheme.get("values") or DEFAULT_SCHEME["values"]
            for key, label, field_key, default_value in self.VALUE_OPTIONS:
                items.append({
                    "key": key,
                    "label": label,
                    "field_key": field_key,
                    "column": (scheme.get("fields") or DEFAULT_SCHEME["fields"]).get(field_key, ""),
                    "value": values.get(key, default_value),
                    "enabled": True,
                })
        return items or copy_json(DEFAULT_SCHEME["value_rules"])

    def _scheme_review_plans(self, scheme):
        return copy_json(scheme.get("review_plans") or DEFAULT_SCHEME["review_plans"])

    def _fill_company_keywords(self, target_var):
        names = [line.strip() for line in self.company_text.get("1.0", "end").splitlines() if line.strip()]
        if not names:
            messagebox.showinfo("提示", "公司名单为空，无法填入关键词。")
            return
        target_var.set("，".join(names))

    def add_field_row(self, item=None):
        item = item or {"key": "custom_field", "label": "自定义字段", "column": "", "enabled": True}
        row_index = len(self.field_row_vars)
        row = ctk.CTkFrame(self.field_rows_frame, fg_color=SURFACE, corner_radius=8, border_width=1, border_color=BORDER)
        row.grid(row=row_index, column=0, sticky="ew", padx=(0, 4), pady=(0, 8))
        row.columnconfigure(1, weight=1)
        row.columnconfigure(3, weight=1)
        enabled_var = tk.BooleanVar(value=item.get("enabled", True) is not False)
        key_var = tk.StringVar(value=self._field_display(item.get("key") or "custom_field"))
        label_var = tk.StringVar(value=item.get("label") or item.get("key") or "")
        column = item.get("column") or ""
        if isinstance(column, list):
            column = ",".join(str(part) for part in column)
        column_var = tk.StringVar(value=str(column))
        ui_check(row, text="启用", variable=enabled_var).grid(row=0, column=0, sticky="w", padx=(12, 10), pady=(10, 3))
        ui_label(row, text="用途").grid(row=0, column=1, sticky="w", pady=(10, 3))
        ui_label(row, text="显示名称").grid(row=0, column=3, sticky="w", pady=(10, 3))
        ModernComboBox(row, variable=key_var, values=[self._field_display(key) for key, _ in self.FIELD_OPTIONS], state="readonly").grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=3)
        ui_entry(row, textvariable=label_var).grid(row=1, column=3, sticky="ew", padx=(0, 12), pady=3)
        ui_label(row, text="匹配列 / 字段列").grid(row=2, column=0, columnspan=2, sticky="w", padx=(12, 0), pady=(7, 3))
        combo = ModernComboBox(row, variable=column_var, values=self._column_choices() or [""])
        combo.grid(row=3, column=0, columnspan=4, sticky="ew", padx=(12, 12), pady=(3, 12))
        ui_button(row, text="删除字段", command=lambda: self._delete_dynamic_row(self.field_row_vars, row), width=100).grid(row=3, column=4, sticky="e", padx=(0, 12), pady=(3, 12))
        self.field_row_vars.append({"frame": row, "enabled": enabled_var, "key": key_var, "label": label_var, "column": column_var, "combo": combo})

    def add_plan_row(self, item=None):
        item = item or {
            "name": "新质检计划",
            "role": "",
            "enabled": True,
            "output_sheet": True,
            "match_type": "条件筛选",
            "keyword_columns": "",
            "keywords": "",
            "apply_conditions": False,
            "conditions": [{"column": "", "operator": "等于", "value": ""}],
            "sampling": {"enabled": False, "mode": "按比例", "value": 0.2, "min_count": 1},
            "overtime": {"enabled": False, "mode": "按起止时间计算", "send_column": "A", "process_column": "AB", "duration_column": "", "threshold_minutes": 20, "id_column": "B"},
        }
        sampling = item.get("sampling") or {}
        overtime = item.get("overtime") or {}
        row_index = len(self.plan_row_vars)
        card = ctk.CTkFrame(self.plan_rows_frame, fg_color=SURFACE, corner_radius=8, border_width=1, border_color=BORDER)
        card.grid(row=row_index, column=0, sticky="ew", padx=(0, 4), pady=(0, 10))
        for col in range(4):
            card.columnconfigure(col, weight=1)

        enabled_var = tk.BooleanVar(value=item.get("enabled", True) is not False)
        output_var = tk.BooleanVar(value=item.get("output_sheet", True) is not False)
        name_var = tk.StringVar(value=item.get("name") or "质检计划")
        match_type_var = tk.StringVar(value=item.get("match_type") or "条件筛选")
        columns_var = tk.StringVar(value=str(item.get("keyword_columns") or ""))
        keywords_var = tk.StringVar(value=str(item.get("keywords") or ""))
        apply_conditions_var = tk.BooleanVar(value=item.get("apply_conditions") is True)
        conditions_var = tk.StringVar(value=self._conditions_to_text(item.get("conditions") or []))
        sampling_enabled_var = tk.BooleanVar(value=sampling.get("enabled") is True)
        sampling_mode_var = tk.StringVar(value=sampling.get("mode") or "按比例")
        sampling_value_var = tk.StringVar(value=str(sampling.get("value", 0.2)))
        sampling_min_var = tk.StringVar(value=str(sampling.get("min_count", 1)))
        overtime_enabled_var = tk.BooleanVar(value=overtime.get("enabled") is True)
        overtime_mode_var = tk.StringVar(value=overtime.get("mode") or "按起止时间计算")
        send_var = tk.StringVar(value=str(overtime.get("send_column") or "A"))
        process_var = tk.StringVar(value=str(overtime.get("process_column") or "AB"))
        duration_var = tk.StringVar(value=str(overtime.get("duration_column") or ""))
        threshold_var = tk.StringVar(value=str(overtime.get("threshold_minutes", 20)))
        id_var = tk.StringVar(value=str(overtime.get("id_column") or "B"))

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=4, sticky="ew", padx=12, pady=(12, 8))
        header.columnconfigure(1, weight=1)
        ui_check(header, text="启用", variable=enabled_var).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ui_entry(header, textvariable=name_var).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ModernComboBox(header, variable=match_type_var, values=["按月份专项关键词", "关键词筛选", "条件筛选"], state="readonly", width=18).grid(row=0, column=2, sticky="e", padx=(0, 8))
        ui_check(header, text="导出工作表", variable=output_var).grid(row=0, column=3, sticky="e", padx=(0, 8))
        ui_button(header, text="删除计划", command=lambda: self._delete_dynamic_row(self.plan_row_vars, card), width=100).grid(row=0, column=4, sticky="e")

        ui_label(card, text="关键词匹配列").grid(row=1, column=0, sticky="w", padx=(12, 0), pady=(0, 3))
        ui_label(card, text="关键词").grid(row=1, column=1, sticky="w", pady=(0, 3))
        condition_header = ctk.CTkFrame(card, fg_color="transparent")
        condition_header.grid(row=1, column=2, columnspan=2, sticky="ew", padx=(0, 12), pady=(0, 3))
        ui_label(condition_header, text="条件筛选").pack(side="left")
        ui_check(condition_header, text="关键词命中后继续按条件筛选", variable=apply_conditions_var).pack(side="left", padx=(12, 0))
        col_combo = ModernComboBox(card, variable=columns_var, values=self._column_choices() or [""])
        col_combo.grid(row=2, column=0, sticky="ew", padx=(12, 8), pady=3)
        keyword_frame = ctk.CTkFrame(card, fg_color="transparent")
        keyword_frame.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=3)
        keyword_frame.columnconfigure(0, weight=1)
        ui_entry(keyword_frame, textvariable=keywords_var).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ui_button(keyword_frame, text="填入公司名单", command=lambda: self._fill_company_keywords(keywords_var), width=120).grid(row=0, column=1)
        ui_entry(card, textvariable=conditions_var).grid(row=2, column=2, columnspan=2, sticky="ew", padx=(0, 12), pady=3)

        ui_label(card, text="关键词用逗号、顿号或换行分隔；条件示例：A不为空；A=同意,不同意；R=否;S=是。", foreground="#546179").grid(row=3, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 8))

        sample_box = ctk.CTkFrame(card, fg_color="#f7fafc", corner_radius=7, border_width=1, border_color="#e2eaf2")
        sample_box.grid(row=4, column=0, columnspan=4, sticky="ew", padx=12, pady=(0, 8))
        sample_box.columnconfigure(1, weight=1)
        sample_box.columnconfigure(3, weight=1)
        ui_check(sample_box, text="启用随机抽样", variable=sampling_enabled_var).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 5))
        ui_label(sample_box, text="抽样方式").grid(row=1, column=0, sticky="w", padx=(10, 6), pady=3)
        ModernComboBox(sample_box, variable=sampling_mode_var, values=["按比例", "按数量"], state="readonly", width=12).grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=3)
        ui_label(sample_box, text="抽样值").grid(row=1, column=2, sticky="w", padx=(0, 6), pady=3)
        ui_entry(sample_box, textvariable=sampling_value_var, width=12).grid(row=1, column=3, sticky="ew", pady=3)
        ui_label(sample_box, text="最少条数").grid(row=2, column=0, sticky="w", padx=(10, 6), pady=(3, 10))
        ui_entry(sample_box, textvariable=sampling_min_var, width=12).grid(row=2, column=1, sticky="ew", padx=(0, 18), pady=(3, 10))
        ui_label(sample_box, text="按比例：0.2 表示 20%，也可填 20；按数量：抽样值表示固定抽取条数。", foreground="#546179").grid(row=2, column=2, columnspan=2, sticky="w", padx=(0, 10), pady=(3, 10))

        overtime_box = ctk.CTkFrame(card, fg_color="#f7fafc", corner_radius=7, border_width=1, border_color="#e2eaf2")
        overtime_box.grid(row=5, column=0, columnspan=4, sticky="ew", padx=12, pady=(0, 12))
        overtime_box.columnconfigure(1, weight=1)
        overtime_box.columnconfigure(3, weight=1)
        overtime_box.columnconfigure(5, weight=1)
        ui_check(overtime_box, text="启用超时检查", variable=overtime_enabled_var).grid(row=0, column=0, columnspan=6, sticky="w", padx=10, pady=(10, 5))
        ui_label(overtime_box, text="计算方式").grid(row=1, column=0, sticky="w", padx=(10, 6), pady=3)
        ModernComboBox(overtime_box, variable=overtime_mode_var, values=["按起止时间计算", "使用已有处理时长列"], state="readonly", width=18).grid(row=1, column=1, sticky="ew", padx=(0, 14), pady=3)
        ui_label(overtime_box, text="阈值(分钟)").grid(row=1, column=2, sticky="w", padx=(0, 6), pady=3)
        ui_entry(overtime_box, textvariable=threshold_var, width=12).grid(row=1, column=3, sticky="ew", padx=(0, 14), pady=3)
        ui_label(overtime_box, text="编号列").grid(row=1, column=4, sticky="w", padx=(0, 6), pady=3)
        id_combo = ModernComboBox(overtime_box, variable=id_var, values=self._column_choices() or [""], width=16)
        id_combo.grid(row=1, column=5, sticky="ew", pady=3)

        ui_label(overtime_box, text="派发时间列").grid(row=2, column=0, sticky="w", padx=(10, 6), pady=3)
        send_combo = ModernComboBox(overtime_box, variable=send_var, values=self._column_choices() or [""], width=16)
        send_combo.grid(row=2, column=1, sticky="ew", padx=(0, 14), pady=3)
        ui_label(overtime_box, text="处理时间列").grid(row=2, column=2, sticky="w", padx=(0, 6), pady=3)
        process_combo = ModernComboBox(overtime_box, variable=process_var, values=self._column_choices() or [""], width=16)
        process_combo.grid(row=2, column=3, sticky="ew", padx=(0, 14), pady=3)
        ui_label(overtime_box, text="已有时长列").grid(row=2, column=4, sticky="w", padx=(0, 6), pady=3)
        duration_combo = ModernComboBox(overtime_box, variable=duration_var, values=self._column_choices() or [""], width=16)
        duration_combo.grid(row=2, column=5, sticky="ew", pady=3)
        ui_label(overtime_box, text="按起止时间计算时使用“处理时间列 - 派发时间列”；如果源表已有处理时长，则选择“使用已有处理时长列”并填写已有时长列。", foreground="#546179").grid(row=3, column=0, columnspan=6, sticky="w", padx=10, pady=(2, 10))

        self.plan_row_vars.append({
            "frame": card,
            "role": item.get("role", ""),
            "enabled": enabled_var,
            "output": output_var,
            "name": name_var,
            "match_type": match_type_var,
            "columns": columns_var,
            "keywords": keywords_var,
            "apply_conditions": apply_conditions_var,
            "conditions": conditions_var,
            "sampling_enabled": sampling_enabled_var,
            "sampling_mode": sampling_mode_var,
            "sampling_value": sampling_value_var,
            "sampling_min": sampling_min_var,
            "overtime_enabled": overtime_enabled_var,
            "overtime_mode": overtime_mode_var,
            "send": send_var,
            "process": process_var,
            "duration": duration_var,
            "threshold": threshold_var,
            "id": id_var,
            "combos": [col_combo, send_combo, process_combo, duration_combo, id_combo],
        })

    def _delete_dynamic_row(self, rows, frame):
        for idx, item in enumerate(list(rows)):
            if item["frame"] == frame:
                item["frame"].destroy()
                del rows[idx]
                break

    def _refresh_column_choices(self):
        for item in self.field_row_vars + self.value_row_vars:
            item["combo"]["values"] = self._column_choices()
        for item in self.plan_row_vars:
            for combo in item["combos"]:
                combo["values"] = self._column_choices()

    def _collect_field_items(self):
        items = []
        for row in self.field_row_vars:
            key = self._parse_field_key(row["key"].get())
            label = row["label"].get().strip() or key
            column = row["column"].get().strip()
            if not key and not column:
                continue
            items.append({"key": key, "label": label, "column": column, "enabled": bool(row["enabled"].get())})
        return items

    def _collect_review_plans(self):
        plans = []
        for row in self.plan_row_vars:
            try:
                sampling_value = float(row["sampling_value"].get() or 0)
                sampling_min = int(row["sampling_min"].get() or 1)
                threshold = float(row["threshold"].get() or 20)
            except ValueError:
                messagebox.showerror("配置错误", "抽样值、最少条数和超时阈值必须是数字。")
                return None
            plans.append({
                "name": row["name"].get().strip() or "质检计划",
                "role": row.get("role", ""),
                "enabled": bool(row["enabled"].get()),
                "output_sheet": bool(row["output"].get()),
                "match_type": row["match_type"].get().strip() or "条件筛选",
                "keyword_columns": row["columns"].get().strip(),
                "keywords": row["keywords"].get().strip(),
                "match_mode": "包含",
                "case_sensitive": False,
                "apply_conditions": bool(row["apply_conditions"].get()),
                "conditions": self._parse_conditions(row["conditions"].get()),
                "sampling": {
                    "enabled": bool(row["sampling_enabled"].get()),
                    "mode": row["sampling_mode"].get().strip() or "按比例",
                    "value": sampling_value,
                    "min_count": sampling_min,
                },
                "overtime": {
                    "enabled": bool(row["overtime_enabled"].get()),
                    "mode": row["overtime_mode"].get().strip() or "按起止时间计算",
                    "send_column": row["send"].get().strip() or "A",
                    "process_column": row["process"].get().strip() or "AB",
                    "duration_column": row["duration"].get().strip(),
                    "threshold_minutes": threshold,
                    "id_column": row["id"].get().strip() or "B",
                },
            })
        return plans

    def _validate_scheme_for_save(self, scheme):
        errors = []
        if not (scheme.get("name") or "").strip():
            errors.append("方案名称不能为空。")
        plans = [plan for plan in scheme.get("review_plans") or [] if plan.get("enabled", True)]
        if not plans:
            errors.append("至少需要启用一个质检计划。")
        for idx, plan in enumerate(plans, start=1):
            name = plan.get("name") or f"质检计划{idx}"
            match_type = plan.get("match_type", "条件筛选")
            if match_type in {"关键词筛选", "按月份专项关键词"} and not plan.get("keyword_columns"):
                errors.append(f"质检计划【{name}】未配置关键词匹配列。")
            if match_type == "关键词筛选" and not plan.get("keywords"):
                errors.append(f"质检计划【{name}】选择了关键词筛选，但关键词为空。")
            if (match_type == "条件筛选" or plan.get("apply_conditions")) and not plan.get("conditions"):
                errors.append(f"质检计划【{name}】需要条件筛选，但条件为空。")
            sampling = plan.get("sampling") or {}
            if sampling.get("enabled"):
                if float(sampling.get("value", 0) or 0) <= 0:
                    errors.append(f"质检计划【{name}】抽样值必须大于 0。")
                if int(sampling.get("min_count", 0) or 0) < 0:
                    errors.append(f"质检计划【{name}】最少条数不能小于 0。")
            overtime = plan.get("overtime") or {}
            if overtime.get("enabled"):
                if not overtime.get("id_column"):
                    errors.append(f"质检计划【{name}】启用超时检查后必须配置编号列。")
                if float(overtime.get("threshold_minutes", 20) or 20) < 0:
                    errors.append(f"质检计划【{name}】超时阈值不能小于 0。")
                if overtime.get("mode") == "使用已有处理时长列":
                    if not overtime.get("duration_column"):
                        errors.append(f"质检计划【{name}】使用已有处理时长列时，必须配置已有时长列。")
                elif not overtime.get("send_column") or not overtime.get("process_column"):
                    errors.append(f"质检计划【{name}】按起止时间计算时，必须配置派发时间列和处理时间列。")
        return errors

    def _value_rules_from_plans(self, plans):
        values = {}
        columns = {}
        for plan in plans:
            for condition in plan.get("conditions") or []:
                name = plan.get("name", "")
                value = condition.get("value", "")
                column = condition.get("column", "")
                if "无效" in name:
                    values["valid_no"] = value
                    columns["valid_scope"] = column
                elif "舆情提醒" in name:
                    values["reminder"] = value
                    columns["reminder"] = column
        rules = copy_json(DEFAULT_SCHEME["value_rules"])
        for rule in rules:
            if rule["key"] in values:
                rule["value"] = values[rule["key"]]
            if rule.get("field_key") in columns:
                rule["column"] = columns[rule["field_key"]]
        return rules

    @staticmethod
    def _conditions_to_text(conditions):
        parts = []
        for condition in conditions:
            column = str(condition.get("column") or "").strip()
            value = str(condition.get("value") or "").strip()
            operator = str(condition.get("operator") or "等于").strip()
            if not column and not value:
                continue
            if operator == "等于":
                parts.append(f"{column}={value}")
            elif operator in {"非空", "不为空", "为空"}:
                parts.append(f"{column}{operator}")
            else:
                parts.append(f"{column}{operator}{value}")
        return ";".join(parts)

    @staticmethod
    def _parse_conditions(text):
        conditions = []
        for part in str(text or "").replace("；", ";").split(";"):
            item = part.strip()
            if not item:
                continue
            operator = "等于"
            if item.endswith("不为空"):
                column, value = item[:-3], ""
                operator = "不为空"
            elif item.endswith("非空"):
                column, value = item[:-2], ""
                operator = "非空"
            elif item.endswith("为空"):
                column, value = item[:-2], ""
                operator = "为空"
            elif "!=" in item:
                column, value = item.split("!=", 1)
                operator = "不等于"
            elif "=" in item:
                column, value = item.split("=", 1)
            elif "不包含" in item:
                column, value = item.split("不包含", 1)
                operator = "不包含"
            elif "包含" in item:
                column, value = item.split("包含", 1)
                operator = "包含"
            else:
                column, value = item, ""
            conditions.append({"column": column.strip(), "operator": operator, "value": value.strip()})
        return conditions

    def _collect_value_rules(self):
        field_map = {key: field_key for key, _label, field_key, _default in self.VALUE_OPTIONS}
        items = []
        for row in self.value_row_vars:
            key = self._parse_key(row["key"].get())
            label = row["label"].get().strip() or key
            column = row["column"].get().strip()
            value = row["value"].get().strip()
            if not key and not value:
                continue
            items.append({
                "key": key,
                "label": label,
                "field_key": field_map.get(key, ""),
                "column": column,
                "value": value,
                "enabled": bool(row["enabled"].get()),
            })
        return items


def load_config_headers(path):
    row, _max_row, _max_col = load_workbook_headers(path)
    choices = []
    for idx, value in enumerate(row, start=1):
        text = str(value).strip() if value is not None else ""
        letter = get_column_name(idx)
        choices.append(f"{letter} {text}" if text else letter)
    return choices


def get_column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


class SpecialConfigWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("专项质检配置")
        self.geometry("900x640")
        self.minsize(820, 560)
        self.transient(parent)
        apply_theme(self)

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

        left = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=10, border_width=1, border_color=BORDER)
        left.grid(row=0, column=0, sticky="ns", padx=(14, 7), pady=14)
        ui_label(left, text="月份", style="Section.TLabel").pack(anchor="w", padx=14, pady=(14, 0))
        month_shell = ctk.CTkFrame(left, fg_color=SURFACE, corner_radius=8, border_width=1, border_color=BORDER)
        month_shell.pack(fill="both", expand=True, padx=14, pady=(8, 14))
        self.month_list = ui_listbox(month_shell, height=12, exportselection=False, width=12)
        self.month_list.pack(fill="both", expand=True, padx=8, pady=8)
        for month in range(1, 13):
            self.month_list.insert("end", f"{month}月")
        self.month_list.bind("<<ListboxSelect>>", self.on_month_select)

        right = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=10, border_width=1, border_color=BORDER)
        right.grid(row=0, column=1, sticky="nsew", padx=(7, 14), pady=14)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(5, weight=1)

        ui_label(right, text="策略名称").grid(row=0, column=0, sticky="w", padx=16, pady=(16, 0))
        ui_entry(right, textvariable=self.name_var).grid(row=1, column=0, sticky="ew", padx=16, pady=(6, 12))

        ui_label(right, text="质检策略").grid(row=2, column=0, sticky="w", padx=16)
        self.strategy_text = ctk.CTkTextbox(right, height=110, wrap="word", corner_radius=8, border_width=1, border_color=BORDER, fg_color="#fbfdff", text_color=TEXT)
        self.strategy_text.grid(row=3, column=0, sticky="ew", padx=16, pady=(6, 12))

        ui_label(right, text="质检关键词（用逗号、顿号或换行分隔）").grid(row=4, column=0, sticky="w", padx=16)
        self.keyword_text = ctk.CTkTextbox(right, height=260, wrap="word", corner_radius=8, border_width=1, border_color=BORDER, fg_color="#fbfdff", text_color=TEXT)
        self.keyword_text.grid(row=5, column=0, sticky="nsew", padx=16, pady=(6, 12))

        actions = ctk.CTkFrame(right, fg_color="transparent")
        actions.grid(row=6, column=0, sticky="ew", padx=16, pady=(0, 16))
        ui_button(actions, text="保存当前月份", command=self.save_current_month, width=130).pack(side="left")
        ui_button(actions, text="恢复当前月份默认值", command=self.reset_current_month, width=170).pack(side="left", padx=8)
        ui_button(actions, text="保存全部配置", command=self.save_all, primary=True, width=140).pack(side="right")

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
