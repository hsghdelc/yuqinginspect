import random
import json
import sys
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from monthly_special import get_monthly_plan


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
    }
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as f:
        loaded = json.load(f)
    default.update({k: v for k, v in loaded.items() if v is not None})
    return default


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


def _copy_row_style(src_ws, dst_ws, src_row, dst_row, max_col):
    for col in range(1, max_col + 1):
        src = src_ws.cell(src_row, col)
        dst = dst_ws.cell(dst_row, col)
        dst.value = src.value
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)


def _copy_column_widths(src_ws, dst_ws, max_col):
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        dst_ws.column_dimensions[letter].width = src_ws.column_dimensions[letter].width


def _write_rows_from_source(src_ws, dst_ws, row_numbers, max_col):
    if row_numbers:
        for out_row, src_row in enumerate(row_numbers, start=1):
            _copy_row_style(src_ws, dst_ws, src_row, out_row, max_col)
    else:
        _copy_row_style(src_ws, dst_ws, 1, 1, max_col)
    _copy_column_widths(src_ws, dst_ws, max_col)
    for row in dst_ws.iter_rows():
        dst_ws.row_dimensions[row[0].row].height = 15


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


def _add_inspection_info(ws, inspector):
    ws.cell(1, 34).value = "质检人员"
    ws.cell(1, 35).value = "质检结果"
    ws.cell(1, 36).value = "补录单号"
    for col in range(34, 37):
        cell = ws.cell(1, col)
        cell.font = copy(ws.cell(1, 1).font)
        cell.font = copy(cell.font)
        cell.font = cell.font.copy(bold=True)
        cell.fill = copy(ws.cell(1, 1).fill)
        ws.column_dimensions[get_column_letter(col)].width = 12 if col < 36 else 15
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 34).value = inspector
        ws.cell(row, 35).value = "通过"
        ws.cell(row, 36).value = ""


def _filter_rows(ws, predicate):
    rows = [1]
    for row in range(2, ws.max_row + 1):
        if predicate(row):
            rows.append(row)
    return rows


def _count(ws, col, value):
    return sum(1 for row in range(2, ws.max_row + 1) if _cell_text(ws.cell(row, col).value) == value)


def _countifs(ws, pairs):
    total = 0
    for row in range(2, ws.max_row + 1):
        if all(_cell_text(ws.cell(row, col).value) == value for col, value in pairs):
            total += 1
    return total


def process_file(input_path, output_dir=None, inspector="未命名质检员", run_date=None, config_path=None):
    input_path = Path(input_path)
    output_dir = Path(output_dir) if output_dir else input_path.parent
    run_date = run_date or datetime.now()
    inspector = inspector.strip() or "未命名质检员"
    config = load_config(config_path)
    companies = set(config.get("companies") or COMPANIES)
    special_target_columns = [int(col) for col in config.get("special_target_columns", [3, 11]) if int(col) > 0]
    overtime_threshold = float(config.get("overtime_threshold_minutes", 20) or 20)

    wb = load_workbook(input_path)
    src_ws = wb.worksheets[0]
    max_col = src_ws.max_column

    kept_after_ag = [1]
    for row in range(2, src_ws.max_row + 1):
        if _cell_text(src_ws.cell(row, 33).value):
            kept_after_ag.append(row)

    filtered_rows = [1]
    for row in kept_after_ag[1:]:
        if _cell_text(src_ws.cell(row, 8).value) in companies:
            filtered_rows.append(row)

    work_wb = Workbook()
    ws2 = work_wb.active
    ws2.title = "Sheet2"
    _write_rows_from_source(src_ws, ws2, filtered_rows, max_col)

    headers = [ws2.cell(1, col).value for col in range(1, ws2.max_column + 1)]
    col_send = _find_col(headers, "工单派发时间", 1)
    col_process = _find_col(headers, "客服部处理时间", 28)
    col_r = _find_col(headers, "是否符合舆情范围", 18)
    col_s = _find_col(headers, "是否营销类舆情事件", 19)
    col_u = _find_col(headers, "是否为重复事件", 21)
    col_id = _find_col(headers, "编号", 2)

    duration_col = ws2.max_column + 1
    ws2.cell(1, duration_col).value = "处理时长(分钟)"
    ws2.column_dimensions[get_column_letter(duration_col)].width = 15
    overtime_ids = []
    for row in range(2, ws2.max_row + 1):
        if (
            _cell_text(ws2.cell(row, col_r).value) == "是"
            and _cell_text(ws2.cell(row, col_s).value) == "是"
            and _cell_text(ws2.cell(row, col_u).value) == "否"
        ):
            send_time = _parse_datetime(ws2.cell(row, col_send).value)
            process_time = _parse_datetime(ws2.cell(row, col_process).value)
            if send_time and process_time:
                minutes = round((process_time - send_time).total_seconds() / 60, 2)
                ws2.cell(row, duration_col).value = minutes
                if minutes > overtime_threshold:
                    overtime_ids.append(_cell_text(ws2.cell(row, col_id).value))
            else:
                ws2.cell(row, duration_col).value = "时间格式错误"

    month_plan = get_monthly_plan(run_date.month)
    special_sheet_name = _safe_sheet_name(month_plan["name"] + "复核")
    keywords = [kw.lower() for kw in month_plan["keywords"]]

    ws_special = work_wb.create_sheet(special_sheet_name)
    special_rows = [1]
    for row in range(2, ws2.max_row + 1):
        text = " ".join(_cell_text(ws2.cell(row, col).value) for col in special_target_columns).lower()
        if any(keyword in text for keyword in keywords):
            special_rows.append(row)
    _write_rows_from_source(ws2, ws_special, special_rows, ws2.max_column)

    ws_non_marketing = work_wb.create_sheet("Sheet4非营销")
    non_marketing_rows = _filter_rows(
        ws2,
        lambda row: _cell_text(ws2.cell(row, 15).value) == "负面事件"
        and _cell_text(ws2.cell(row, 18).value) == "否"
        and _cell_text(ws2.cell(row, 20).value) != "舆情提醒",
    )
    _write_rows_from_source(ws2, ws_non_marketing, non_marketing_rows, ws2.max_column)

    ws_reminder = work_wb.create_sheet("Sheet5舆情提醒复核")
    reminder_rows = _filter_rows(ws2, lambda row: _cell_text(ws2.cell(row, 20).value) == "舆情提醒")
    _write_rows_from_source(ws2, ws_reminder, reminder_rows, ws2.max_column)

    ws_invalid = work_wb.create_sheet("Sheet6无效复核")
    invalid_rows = _filter_rows(ws2, lambda row: _cell_text(ws2.cell(row, 18).value) == "否")
    if len(invalid_rows) > 1:
        data_rows = invalid_rows[1:]
        keep_count = max(1, int((len(data_rows) * 0.2) + 0.999999))
        invalid_rows = [1] + random.sample(data_rows, min(keep_count, len(data_rows)))
    _write_rows_from_source(ws2, ws_invalid, invalid_rows, ws2.max_column)

    for sheet in (ws_reminder, ws_invalid, ws_special):
        _add_inspection_info(sheet, inspector)

    a = sum(1 for row in range(2, ws2.max_row + 1) if _cell_text(ws2.cell(row, 1).value))
    b = _count(ws2, 18, "否")
    c = _count(ws2, 15, "正面事件")
    d = _countifs(ws2, [(15, "负面事件"), (19, "否")])
    e = _countifs(ws2, [(19, "是"), (18, "是")])
    f = _countifs(ws2, [(19, "是"), (18, "是"), (21, "是")])
    g = _count(ws2, 20, "舆情提醒")
    h = _countifs(ws2, [(25, "民生类舆情"), (18, "是"), (19, "是")])

    report_text = (
        f"{run_date.year}年{run_date.month}月{run_date.day}日，南方分中心共收到舆情工单待办{a}件，"
        f"其中无效事件{b}件、正面事件{c}件、负面非营销类事件{d}件，舆情提醒{g}件，"
        f"负面营销类舆情{e}件（重复事件{f}件，民生事件{h}件）。"
    )

    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    for source_sheet, target_name in (
        (ws_special, ws_special.title),
        (ws_invalid, "无效复核"),
        (ws_reminder, "舆情提醒复核"),
    ):
        target = out_wb.create_sheet(target_name)
        _write_rows_from_source(source_sheet, target, list(range(1, source_sheet.max_row + 1)), source_sheet.max_column)

    yesterday = run_date - timedelta(days=1)
    file_name = f"{yesterday.year}年{yesterday.month}月{yesterday.day}日8点-{run_date.year}年{run_date.month}月{run_date.day}日8点舆情质检明细({inspector}).xlsx"
    output_path = output_dir / file_name
    out_wb.save(output_path)

    return {
        "output_path": str(output_path),
        "report_text": report_text,
        "special_name": month_plan["name"],
        "special_count": max(0, ws_special.max_row - 1),
        "reminder_count": max(0, ws_reminder.max_row - 1),
        "invalid_count": max(0, ws_invalid.max_row - 1),
        "overtime_count": len(overtime_ids),
        "overtime_ids": overtime_ids,
    }
