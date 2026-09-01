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


DEFAULT_SCHEME_ID = "default"

DEFAULT_SCHEME = {
    "id": DEFAULT_SCHEME_ID,
    "name": "默认宏结构方案",
    "fields": {
        "keep": "AG",
        "company": "H",
        "send_time": "A",
        "process_time": "AB",
        "event_nature": "O",
        "valid_scope": "R",
        "marketing": "S",
        "reminder": "T",
        "duplicate": "U",
        "event_category": "Y",
        "id": "B",
        "special_targets": ["C", "K"],
    },
    "field_items": [
        {"key": "keep", "label": "基础保留列", "column": "AG", "enabled": True},
        {"key": "company", "label": "公司筛选列", "column": "H", "enabled": True},
        {"key": "send_time", "label": "工单派发时间列", "column": "A", "enabled": True},
        {"key": "process_time", "label": "客服部处理时间列", "column": "AB", "enabled": True},
        {"key": "event_nature", "label": "事件性质列", "column": "O", "enabled": True},
        {"key": "valid_scope", "label": "是否符合舆情范围列", "column": "R", "enabled": True},
        {"key": "marketing", "label": "是否营销类舆情事件列", "column": "S", "enabled": True},
        {"key": "reminder", "label": "舆情提醒标识列", "column": "T", "enabled": True},
        {"key": "duplicate", "label": "是否重复事件列", "column": "U", "enabled": True},
        {"key": "event_category", "label": "事件分类列", "column": "Y", "enabled": True},
        {"key": "id", "label": "编号列", "column": "B", "enabled": True},
        {"key": "special_targets", "label": "专项关键词匹配列", "column": "C,K", "enabled": True},
    ],
    "values": {
        "company_names": sorted(COMPANIES),
        "valid_yes": "是",
        "valid_no": "否",
        "marketing_yes": "是",
        "marketing_no": "否",
        "duplicate_yes": "是",
        "duplicate_no": "否",
        "negative_event": "负面事件",
        "positive_event": "正面事件",
        "reminder": "舆情提醒",
        "livelihood_event": "民生类舆情",
    },
    "value_rules": [
        {"key": "valid_yes", "label": "符合舆情范围：是", "field_key": "valid_scope", "column": "R", "value": "是", "enabled": True},
        {"key": "valid_no", "label": "符合舆情范围：否", "field_key": "valid_scope", "column": "R", "value": "否", "enabled": True},
        {"key": "marketing_yes", "label": "营销类：是", "field_key": "marketing", "column": "S", "value": "是", "enabled": True},
        {"key": "marketing_no", "label": "营销类：否", "field_key": "marketing", "column": "S", "value": "否", "enabled": True},
        {"key": "duplicate_yes", "label": "重复事件：是", "field_key": "duplicate", "column": "U", "value": "是", "enabled": True},
        {"key": "duplicate_no", "label": "重复事件：否", "field_key": "duplicate", "column": "U", "value": "否", "enabled": True},
        {"key": "negative_event", "label": "负面事件值", "field_key": "event_nature", "column": "O", "value": "负面事件", "enabled": True},
        {"key": "positive_event", "label": "正面事件值", "field_key": "event_nature", "column": "O", "value": "正面事件", "enabled": True},
        {"key": "reminder", "label": "舆情提醒值", "field_key": "reminder", "column": "T", "value": "舆情提醒", "enabled": True},
        {"key": "livelihood_event", "label": "民生事件值", "field_key": "event_category", "column": "Y", "value": "民生类舆情", "enabled": True},
    ],
    "invalid_sample_rate": 0.2,
    "invalid_sample_min": 1,
    "overtime_threshold_minutes": 20,
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
        "active_scheme_id": DEFAULT_SCHEME_ID,
        "schemes": {DEFAULT_SCHEME_ID: DEFAULT_SCHEME},
    }
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as f:
        loaded = json.load(f)
    default.update({k: v for k, v in loaded.items() if v is not None})
    default["schemes"] = _normalize_schemes(default)
    return default


def save_config(config, config_path=None):
    path = Path(config_path) if config_path else _app_dir() / "config.json"
    with path.open("w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def _normalize_schemes(config):
    schemes = copy_json(config.get("schemes")) if config.get("schemes") else {}
    if not isinstance(schemes, dict):
        schemes = {}
    if DEFAULT_SCHEME_ID not in schemes:
        schemes[DEFAULT_SCHEME_ID] = copy_json(DEFAULT_SCHEME)
    for scheme_id, scheme in list(schemes.items()):
        merged = copy_json(DEFAULT_SCHEME)
        merged.update(scheme or {})
        merged["id"] = scheme_id
        merged["fields"] = {**DEFAULT_SCHEME["fields"], **((scheme or {}).get("fields") or {})}
        merged["values"] = {**DEFAULT_SCHEME["values"], **((scheme or {}).get("values") or {})}
        merged["field_items"] = _normalize_field_items(merged)
        merged["value_rules"] = _normalize_value_rules(merged)
        schemes[scheme_id] = merged
    return schemes


def _normalize_field_items(scheme):
    field_items = scheme.get("field_items")
    if not isinstance(field_items, list) or not field_items:
        field_items = []
        fields = scheme.get("fields") or {}
        labels = {item["key"]: item["label"] for item in DEFAULT_SCHEME["field_items"]}
        for key, column in fields.items():
            if isinstance(column, list):
                column = ",".join(str(item) for item in column)
            field_items.append({
                "key": key,
                "label": labels.get(key, key),
                "column": str(column),
                "enabled": True,
            })
    known = {item.get("key"): item for item in field_items if isinstance(item, dict)}
    for default_item in DEFAULT_SCHEME["field_items"]:
        if default_item["key"] not in known:
            field_items.append(copy_json(default_item))
    return [
        {
            "key": str(item.get("key") or "").strip(),
            "label": str(item.get("label") or item.get("key") or "").strip(),
            "column": str(item.get("column") or "").strip(),
            "enabled": item.get("enabled", True) is not False,
        }
        for item in field_items
        if isinstance(item, dict) and str(item.get("key") or "").strip()
    ]


def _normalize_value_rules(scheme):
    value_rules = scheme.get("value_rules")
    if not isinstance(value_rules, list) or not value_rules:
        value_rules = []
        values = scheme.get("values") or {}
        fields = _fields_from_items(scheme.get("field_items") or DEFAULT_SCHEME["field_items"])
        defaults = {item["key"]: item for item in DEFAULT_SCHEME["value_rules"]}
        for key, default_rule in defaults.items():
            rule = copy_json(default_rule)
            rule["value"] = values.get(key, default_rule["value"])
            rule["column"] = fields.get(default_rule["field_key"], default_rule["column"])
            value_rules.append(rule)
    known = {item.get("key"): item for item in value_rules if isinstance(item, dict)}
    for default_rule in DEFAULT_SCHEME["value_rules"]:
        if default_rule["key"] not in known:
            value_rules.append(copy_json(default_rule))
    return [
        {
            "key": str(item.get("key") or "").strip(),
            "label": str(item.get("label") or item.get("key") or "").strip(),
            "field_key": str(item.get("field_key") or "").strip(),
            "column": str(item.get("column") or "").strip(),
            "value": str(item.get("value") or "").strip(),
            "enabled": item.get("enabled", True) is not False,
        }
        for item in value_rules
        if isinstance(item, dict) and str(item.get("key") or "").strip()
    ]


def _fields_from_items(field_items):
    fields = {}
    for item in field_items or []:
        if not isinstance(item, dict) or item.get("enabled", True) is False:
            continue
        key = str(item.get("key") or "").strip()
        column = str(item.get("column") or "").strip()
        if not key or not column:
            continue
        if key == "special_targets":
            fields[key] = [part.strip() for part in column.replace("，", ",").replace("、", ",").split(",") if part.strip()]
        else:
            fields[key] = column
    return fields


def _values_from_rules(value_rules):
    values = {}
    field_overrides = {}
    for rule in value_rules or []:
        if not isinstance(rule, dict) or rule.get("enabled", True) is False:
            continue
        key = str(rule.get("key") or "").strip()
        if not key:
            continue
        values[key] = str(rule.get("value") or "").strip()
        field_key = str(rule.get("field_key") or "").strip()
        column = str(rule.get("column") or "").strip()
        if field_key and column:
            field_overrides[field_key] = column
    return values, field_overrides


def copy_json(value):
    return json.loads(json.dumps(value, ensure_ascii=False))


def get_active_scheme(config, scheme_id=None):
    schemes = _normalize_schemes(config)
    selected_id = scheme_id or config.get("active_scheme_id") or DEFAULT_SCHEME_ID
    if selected_id not in schemes:
        selected_id = DEFAULT_SCHEME_ID
    return schemes[selected_id]


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


def _letters_to_col(text):
    value = 0
    for ch in text.upper():
        if not ("A" <= ch <= "Z"):
            return None
        value = value * 26 + ord(ch) - ord("A") + 1
    return value or None


def _resolve_col(spec, headers, default_index):
    if spec is None or spec == "":
        return default_index
    text = str(spec).strip()
    if not text:
        return default_index
    first_token = text.split()[0] if text.split() else text
    token_letter_col = _letters_to_col(first_token)
    if token_letter_col:
        return token_letter_col
    if text.isdigit():
        return int(text)
    letter_col = _letters_to_col(text)
    if letter_col:
        return letter_col
    for idx, value in enumerate(headers, start=1):
        if _cell_text(value) == text:
            return idx
    for idx, value in enumerate(headers, start=1):
        if text in _cell_text(value):
            return idx
    return default_index


def _resolve_cols(specs, headers, default_indexes):
    if isinstance(specs, str):
        specs = [item.strip() for item in specs.replace("，", ",").replace("、", ",").split(",") if item.strip()]
    if not specs:
        return default_indexes
    columns = []
    for idx, spec in enumerate(specs):
        default_index = default_indexes[min(idx, len(default_indexes) - 1)]
        col = _resolve_col(spec, headers, default_index)
        if col > 0 and col not in columns:
            columns.append(col)
    return columns or default_indexes


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


def process_file(input_path, output_dir=None, inspector="未命名质检员", run_date=None, config_path=None, progress_callback=None, scheme_id=None):
    def progress(message):
        if progress_callback:
            progress_callback(message)

    input_path = Path(input_path)
    output_dir = Path(output_dir) if output_dir else input_path.parent
    run_date = run_date or datetime.now()
    inspector = inspector.strip() or "未命名质检员"
    config = load_config(config_path)
    scheme = get_active_scheme(config, scheme_id)
    fields = {**(scheme.get("fields") or DEFAULT_SCHEME["fields"]), **_fields_from_items(scheme.get("field_items"))}
    rule_values, field_overrides = _values_from_rules(scheme.get("value_rules"))
    fields.update({key: value for key, value in field_overrides.items() if value})
    values = {**(scheme.get("values") or DEFAULT_SCHEME["values"]), **rule_values}
    companies = set(values.get("company_names") or config.get("companies") or COMPANIES)
    overtime_threshold = float(scheme.get("overtime_threshold_minutes") or config.get("overtime_threshold_minutes", 20) or 20)
    invalid_sample_rate = float(scheme.get("invalid_sample_rate", 0.2) or 0.2)
    invalid_sample_min = int(scheme.get("invalid_sample_min", 1) or 1)

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

    headers = list(header)
    col_keep = _resolve_col(fields.get("keep"), headers, 33)
    col_company = _resolve_col(fields.get("company"), headers, 8)
    col_send = _resolve_col(fields.get("send_time"), headers, _find_col(headers, "工单派发时间", 1))
    col_process = _resolve_col(fields.get("process_time"), headers, _find_col(headers, "客服部处理时间", 28))
    col_event_nature = _resolve_col(fields.get("event_nature"), headers, 15)
    col_r = _resolve_col(fields.get("valid_scope"), headers, _find_col(headers, "是否符合舆情范围", 18))
    col_s = _resolve_col(fields.get("marketing"), headers, _find_col(headers, "是否营销类舆情事件", 19))
    col_reminder = _resolve_col(fields.get("reminder"), headers, 20)
    col_u = _resolve_col(fields.get("duplicate"), headers, _find_col(headers, "是否为重复事件", 21))
    col_category = _resolve_col(fields.get("event_category"), headers, 25)
    col_id = _resolve_col(fields.get("id"), headers, _find_col(headers, "编号", 2))
    special_target_columns = _resolve_cols(fields.get("special_targets") or config.get("special_target_columns"), headers, [3, 11])

    progress(f"当前方案：{scheme.get('name', DEFAULT_SCHEME['name'])}")
    progress(f"专项匹配列：{', '.join(get_column_letter(col) for col in special_target_columns)}")

    filtered_rows = []
    scanned_count = 0
    ag_count = 0
    for row in source_rows:
        scanned_count += 1
        if not _row_text(row, col_keep):
            continue
        ag_count += 1
        if _row_text(row, col_company) in companies:
            filtered_rows.append(tuple(row))
    wb.close()
    progress(f"AG 列非空：{ag_count} 行；公司筛选后保留：{len(filtered_rows)} 行")

    duration_col = max_col + 1
    header = _set_row_value(header, duration_col, "处理时长(分钟)")
    overtime_ids = []
    progress("统计超时舆情...")
    processed_rows = []
    for row in filtered_rows:
        if (
            _row_text(row, col_r) == values.get("valid_yes", "是")
            and _row_text(row, col_s) == values.get("marketing_yes", "是")
            and _row_text(row, col_u) == values.get("duplicate_no", "否")
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
    reminder_rows = [row for row in processed_rows if _row_text(row, col_reminder) == values.get("reminder", "舆情提醒")]
    invalid_source_rows = [row for row in processed_rows if _row_text(row, col_r) == values.get("valid_no", "否")]
    if invalid_source_rows:
        keep_count = max(invalid_sample_min, int((len(invalid_source_rows) * invalid_sample_rate) + 0.999999))
        invalid_rows = random.sample(invalid_source_rows, min(keep_count, len(invalid_source_rows)))
    else:
        invalid_rows = []

    progress("生成日报送文本...")
    a = sum(1 for row in processed_rows if _row_text(row, 1))
    b = _count(processed_rows, col_r, values.get("valid_no", "否"))
    c = _count(processed_rows, col_event_nature, values.get("positive_event", "正面事件"))
    d = _countifs(processed_rows, [(col_event_nature, values.get("negative_event", "负面事件")), (col_s, values.get("marketing_no", "否"))])
    e = _countifs(processed_rows, [(col_s, values.get("marketing_yes", "是")), (col_r, values.get("valid_yes", "是"))])
    f = _countifs(processed_rows, [(col_s, values.get("marketing_yes", "是")), (col_r, values.get("valid_yes", "是")), (col_u, values.get("duplicate_yes", "是"))])
    g = _count(processed_rows, col_reminder, values.get("reminder", "舆情提醒"))
    h = _countifs(processed_rows, [(col_category, values.get("livelihood_event", "民生类舆情")), (col_r, values.get("valid_yes", "是")), (col_s, values.get("marketing_yes", "是"))])

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
