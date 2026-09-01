import random
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from monthly_special import get_default_monthly_plans, get_monthly_plan


COMPANIES = {
    "国网甘肃电力", "国网河北电力", "国网黑龙江电力", "国网湖北电力", "国网江苏电力",
    "国网蒙东电力", "国网山东电力", "国网陕西电力", "国网上海电力", "国网四川电力",
    "国网西藏电力", "国网浙江电力", "国网重庆电力", "国网冀北电力",
}


def _app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_config(config_path=None):
    path = Path(config_path) if config_path else _app_dir() / "config.json"
    default = {
        "companies": sorted(COMPANIES),
        "special_target_columns": [3, 11],
        "overtime_threshold_minutes": 20,
        "preserve_cell_styles": False,
        "monthly_special_plans": get_default_monthly_plans(),
    }
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as f:
        loaded = json.load(f)
    default.update({k: v for k, v in loaded.items() if v is not None})
    return default


def save_config(config, config_path=None):
    path = Path(config_path) if config_path else _app_dir() / "config.json"
    with path.open("w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def _cell_text(value):
    return "" if value is None else str(value).strip()


def _safe_sheet_name(name):
    for ch in ':\\/?*[]':
        name = name.replace(ch, "")
    return name[:31]


def _find_col(headers, contains, default_index):
    for idx, value in enumerate(headers, start=1):
        if contains in _cell_text(value):
            return idx
    return default_index


def _row_value(row, col):
    idx = col - 1
    if idx < 0 or idx >= len(row):
        return None
    return row[idx]


def _row_text(row, col):
    return _cell_text(_row_value(row, col))


def _set_row_value(row, col, value):
    values = list(row)
    while len(values) < col:
        values.append(None)
    values[col - 1] = value
    return tuple(values)


def _parse_datetime(value):
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text.replace("/", "-"))
    except ValueError:
        return None


def _append_review_sheet(wb, title, header, rows, inspector):
    ws = wb.create_sheet(title)
    header = _set_row_value(header, 34, "质检人员")
    header = _set_row_value(header, 35, "质检结果")
    header = _set_row_value(header, 36, "补录单号")
    ws.append(header)
    for row in rows:
        row = _set_row_value(row, 34, inspector)
        row = _set_row_value(row, 35, "通过")
        row = _set_row_value(row, 36, "")
        ws.append(row)
    for col in range(1, max(36, len(header)) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 15


def _count(rows, col, value):
    return sum(1 for row in rows if _row_text(row, col) == value)


def _countifs(rows, pairs):
    return sum(1 for row in rows if all(_row_text(row, col) == value for col, value in pairs))


def process_file(input_path, output_dir=None, inspector="未命名质检员", run_date=None, config_path=None, progress_callback=None):
    def progress(message):
        if progress_callback:
            progress_callback(message)

    input_path = Path(input_path)
    output_dir = Path(output_dir) if output_dir else input_path.parent
    run_date = run_date or datetime.now()
    inspector = inspector.strip() or "未命名质检员"
    config = load_config(config_path)
    companies = set(config.get("companies") or COMPANIES)
    special_target_columns = [int(col) for col in config.get("special_target_columns", [3, 11]) if int(col) > 0]
    overtime_threshold = float(config.get("overtime_threshold_minutes", 20) or 20)

    progress("读取 Excel 文件...")
    wb = load_workbook(input_path, read_only=True, data_only=False)
    src_ws = wb.worksheets[0]
    max_col = src_ws.max_column

    progress(f"源表数据：{src_ws.max_row - 1} 行，{max_col} 列")
    progress("读取并筛选数据...")
    source_rows = src_ws.iter_rows(values_only=True)
    try:
        header = tuple(next(source_rows))
    except StopIteration:
        raise ValueError("源文件没有可处理的数据")

    filtered_rows = []
    scanned_count = 0
    ag_count = 0
    for row in source_rows:
        scanned_count += 1
        if not _row_text(row, 33):
            continue
        ag_count += 1
        if _row_text(row, 8) in companies:
            filtered_rows.append(tuple(row))
    wb.close()
    progress(f"AG 列非空：{ag_count} 行；公司筛选后保留：{len(filtered_rows)} 行")

    headers = list(header)
    col_send = _find_col(headers, "工单派发时间", 1)
    col_process = _find_col(headers, "客服部处理时间", 28)
    col_r = _find_col(headers, "是否符合舆情范围", 18)
    col_s = _find_col(headers, "是否营销类舆情事件", 19)
    col_u = _find_col(headers, "是否为重复事件", 21)
    col_id = _find_col(headers, "编号", 2)

    duration_col = max_col + 1
    header = _set_row_value(header, duration_col, "处理时长(分钟)")
    overtime_ids = []
    progress("统计超时舆情...")
    processed_rows = []
    for row in filtered_rows:
        if (
            _row_text(row, col_r) == "是"
            and _row_text(row, col_s) == "是"
            and _row_text(row, col_u) == "否"
        ):
            send_time = _parse_datetime(_row_value(row, col_send))
            process_time = _parse_datetime(_row_value(row, col_process))
            if send_time and process_time:
                minutes = round((process_time - send_time).total_seconds() / 60, 2)
                row = _set_row_value(row, duration_col, minutes)
                if minutes > overtime_threshold:
                    overtime_ids.append(_row_text(row, col_id))
            else:
                row = _set_row_value(row, duration_col, "时间格式错误")
        processed_rows.append(row)

    month_plan = get_monthly_plan(run_date.month, config.get("monthly_special_plans"))
    special_sheet_name = _safe_sheet_name(month_plan["name"] + "复核")
    keywords = [kw.lower() for kw in month_plan["keywords"]]

    progress(f"执行月度专项：{month_plan['name']}")
    special_rows = []
    for row in processed_rows:
        text = " ".join(_row_text(row, col) for col in special_target_columns).lower()
        if any(keyword in text for keyword in keywords):
            special_rows.append(row)
    progress(f"月度专项命中：{len(special_rows)} 行")

    progress("生成舆情提醒和无效复核数据...")
    reminder_rows = [row for row in processed_rows if _row_text(row, 20) == "舆情提醒"]
    invalid_source_rows = [row for row in processed_rows if _row_text(row, 18) == "否"]
    if invalid_source_rows:
        keep_count = max(1, int((len(invalid_source_rows) * 0.2) + 0.999999))
        invalid_rows = random.sample(invalid_source_rows, min(keep_count, len(invalid_source_rows)))
    else:
        invalid_rows = []

    progress("生成日报送文本...")
    a = sum(1 for row in processed_rows if _row_text(row, 1))
    b = _count(processed_rows, 18, "否")
    c = _count(processed_rows, 15, "正面事件")
    d = _countifs(processed_rows, [(15, "负面事件"), (19, "否")])
    e = _countifs(processed_rows, [(19, "是"), (18, "是")])
    f = _countifs(processed_rows, [(19, "是"), (18, "是"), (21, "是")])
    g = _count(processed_rows, 20, "舆情提醒")
    h = _countifs(processed_rows, [(25, "民生类舆情"), (18, "是"), (19, "是")])

    report_text = (
        f"{run_date.year}年{run_date.month}月{run_date.day}日，南方分中心共收到舆情工单待办{a}件，"
        f"其中无效事件{b}件、正面事件{c}件、负面非营销类事件{d}件，舆情提醒{g}件，"
        f"负面营销类舆情{e}件（重复事件{f}件，民生事件{h}件）。"
    )

    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    progress("导出最终质检明细...")
    _append_review_sheet(out_wb, special_sheet_name, header, special_rows, inspector)
    _append_review_sheet(out_wb, "无效复核", header, invalid_rows, inspector)
    _append_review_sheet(out_wb, "舆情提醒复核", header, reminder_rows, inspector)

    yesterday = run_date - timedelta(days=1)
    file_name = f"{yesterday.year}年{yesterday.month}月{yesterday.day}日8点-{run_date.year}年{run_date.month}月{run_date.day}日8点舆情质检明细({inspector}).xlsx"
    output_path = output_dir / file_name
    out_wb.save(output_path)
    progress("保存完成")

    return {
        "output_path": str(output_path),
        "report_text": report_text,
        "special_name": month_plan["name"],
        "special_count": len(special_rows),
        "reminder_count": len(reminder_rows),
        "invalid_count": len(invalid_rows),
        "overtime_count": len(overtime_ids),
        "overtime_ids": overtime_ids,
    }
