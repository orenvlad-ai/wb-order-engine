import io
from datetime import date, datetime
from typing import Tuple, List, Dict, Any, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment

from pydantic import ValidationError

from engine.models import SkuInput, InTransitItem, Recommendation
from engine.calc import calculate

REQUIRED_INPUT_COLS = ["sku", "stock_ff", "stock_mp", "plan_sales_per_day"]
OPTIONAL_INPUT_COLS = ["safety_stock_ff", "safety_stock_mp"]
REQUIRED_INTRANSIT_COLS = ["sku","qty","eta_cn_msk"]
SETTINGS_SHEET_NAME = "Настройки заказа"
REQUIRED_SETTINGS_COLS = [
    "prod_lead_time_days",
    "lead_time_cn_msk",
    "lead_time_msk_mp",
    "moq_step_default",
    "safety_stock_ff_default",
    "safety_stock_mp_default",
]

class BadTemplateError(Exception): ...

def _ensure_columns(df: pd.DataFrame, required: List[str], sheet: str):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise BadTemplateError(
            f"На листе '{sheet}' отсутствуют обязательные колонки: {', '.join(missing)}."
        )


def _is_blank(value: Any) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _parse_int(value: Any, *, sheet: str, column: str, sku: Optional[str] = None) -> int:
    if _is_blank(value):
        target = f" для SKU '{sku}'" if sku else ""
        raise BadTemplateError(
            f"На листе '{sheet}' колонка '{column}'{target} не заполнена."
        )
    try:
        return int(value)
    except (TypeError, ValueError):
        target = f" для SKU '{sku}'" if sku else ""
        raise BadTemplateError(
            f"На листе '{sheet}' колонка '{column}'{target} должна содержать целое число."
        )


def _parse_float(value: Any, *, sheet: str, column: str, sku: str) -> float:
    if _is_blank(value):
        raise BadTemplateError(
            f"На листе '{sheet}' колонка '{column}' для SKU '{sku}' не заполнена."
        )
    try:
        return float(value)
    except (TypeError, ValueError):
        raise BadTemplateError(
            f"На листе '{sheet}' колонка '{column}' для SKU '{sku}' должна содержать число."
        )


def _read_settings(df_settings: pd.DataFrame) -> Dict[str, int]:
    # Считываем первую заполненную строку с общими параметрами заказа
    _ensure_columns(df_settings, REQUIRED_SETTINGS_COLS, SETTINGS_SHEET_NAME)
    df_clean = df_settings.dropna(how="all")
    if df_clean.empty:
        raise BadTemplateError(
            "Лист 'Настройки заказа' пуст. Заполни строку с параметрами по умолчанию."
        )

    row = df_clean.iloc[0].to_dict()
    settings: Dict[str, int] = {}
    for key in REQUIRED_SETTINGS_COLS:
        value = row.get(key)
        settings[key] = _parse_int(value, sheet=SETTINGS_SHEET_NAME, column=key)
    return settings

def read_input(xlsx_bytes: bytes) -> Tuple[List[SkuInput], List[InTransitItem]]:
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    try:
        df_in = pd.read_excel(xl, "Input")
    except ValueError as e:
        raise BadTemplateError("В файле нет листа 'Input'.") from e

    try:
        df_settings = pd.read_excel(xl, SETTINGS_SHEET_NAME)
    except ValueError as e:
        raise BadTemplateError("В файле нет листа 'Настройки заказа'.") from e

    df_in = df_in.where(pd.notna(df_in), None)
    df_settings = df_settings.where(pd.notna(df_settings), None)

    _ensure_columns(df_in, REQUIRED_INPUT_COLS, "Input")
    for col in OPTIONAL_INPUT_COLS:
        if col not in df_in.columns:
            df_in[col] = None

    settings = _read_settings(df_settings)

    try:
        df_tr = pd.read_excel(xl, "InTransit")
        df_tr = df_tr.where(pd.notna(df_tr), None)
        _ensure_columns(df_tr, REQUIRED_INTRANSIT_COLS, "InTransit")
    except ValueError:
        df_tr = pd.DataFrame(columns=REQUIRED_INTRANSIT_COLS)
    except BadTemplateError:
        raise

    items: List[SkuInput] = []
    moq_step_default = settings["moq_step_default"]
    safety_stock_ff_default = settings["safety_stock_ff_default"]
    safety_stock_mp_default = settings["safety_stock_mp_default"]
    prod_lead_time_days = settings["prod_lead_time_days"]
    lead_time_cn_msk = settings["lead_time_cn_msk"]
    lead_time_msk_mp = settings["lead_time_msk_mp"]

    for r in df_in.to_dict("records"):
        if all(_is_blank(v) for v in r.values()):
            continue
        sku = str(r.get("sku") or "").strip()
        if not sku:
            raise BadTemplateError("На листе 'Input' есть строка без SKU. Удали пустые строки.")
        try:
            stock_ff = _parse_int(r.get("stock_ff"), sheet="Input", column="stock_ff", sku=sku)
            stock_mp = _parse_int(r.get("stock_mp"), sheet="Input", column="stock_mp", sku=sku)
            plan_sales_per_day = _parse_float(
                r.get("plan_sales_per_day"), sheet="Input", column="plan_sales_per_day", sku=sku
            )

            raw_mp = r.get("safety_stock_mp")
            if _is_blank(raw_mp):
                safety_stock_mp = safety_stock_mp_default  # пусто → берём дефолт из настроек
            else:
                safety_stock_mp = _parse_int(raw_mp, sheet="Input", column="safety_stock_mp", sku=sku)

            raw_ff = r.get("safety_stock_ff")
            if _is_blank(raw_ff):
                safety_stock_ff = safety_stock_ff_default  # пусто → берём дефолт из настроек
            else:
                safety_stock_ff = _parse_int(raw_ff, sheet="Input", column="safety_stock_ff", sku=sku)

            items.append(SkuInput(
                sku=sku,
                stock_ff=stock_ff,
                stock_mp=stock_mp,
                plan_sales_per_day=plan_sales_per_day,
                prod_lead_time_days=prod_lead_time_days,
                lead_time_cn_msk=lead_time_cn_msk,
                lead_time_msk_mp=lead_time_msk_mp,
                safety_stock_mp=safety_stock_mp,
                safety_stock_ff=safety_stock_ff,
                moq_step=moq_step_default,
            ))
        except (ValueError, ValidationError) as e:
            raise BadTemplateError(
                f"На листе 'Input' строка для SKU '{sku}' содержит неверные данные."
            ) from e

    trans: List[InTransitItem] = []
    for r in df_tr.to_dict("records"):
        if r.get("sku") is None:
            continue
        try:
            eta = pd.to_datetime(r["eta_cn_msk"]).date() if pd.notna(r["eta_cn_msk"]) else None
            if eta is None:
                raise ValueError("Пустая дата ETA")
            trans.append(InTransitItem(
                sku=str(r["sku"]),
                qty=int(r["qty"]),
                eta_cn_msk=eta,
            ))
        except (ValueError, ValidationError) as e:
            raise BadTemplateError(f"InTransit sku={r.get('sku')} имеет неверные qty/дату.") from e

    return items, trans

# ---------- Форматирование Recommendations ----------

_HEADER_FILL = PatternFill(start_color="FFEFEFEF", end_color="FFEFEFEF", fill_type="solid")
_RISK_FILL   = PatternFill(start_color="FFFFE5E5", end_color="FFFFE5E5", fill_type="solid")
_BOLD = Font(bold=True)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
_THIN = Side(border_style="thin", color="FFBFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# Порядок колонок (берём те, что реально есть в данных)
_ORDER = [
    "sku", "order_qty", "shortage", "target", "coverage", "inbound",
    "demand_H", "H_days", "moq_step", "comment", "algo_version"
]

def _order_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in _ORDER if c in df.columns] + [c for c in df.columns if c not in _ORDER]
    return df[cols]

def _apply_formats(ws):
    # Шапка
    for cell in ws[1]:
        cell.font = _BOLD
        cell.alignment = _CENTER
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
        # краткие подсказки
        if cell.value == "order_qty":
            cell.comment = Comment("Рекомендованный заказ с кратностью MOQ", "WB Engine")
        if cell.value == "shortage":
            cell.comment = Comment("Нехватка к цели (demand_H + safety_stock)", "WB Engine")

    # Форматы чисел (без десятых)
    int_like = {"H_days","inbound","coverage","target","shortage","moq_step","order_qty","demand_H"}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column_letter and ws.cell(row=1, column=cell.column).value in int_like:
                cell.number_format = "0"
            cell.border = _BORDER
            if isinstance(cell.value, str):
                cell.alignment = _LEFT

    # Подсветка риска: если shortage > 0 → розовым фоном
    # (risk_flag у нас = shortage>0, т.к. отдельного поля нет)
    shortage_col = None
    for cell in ws[1]:
        if cell.value == "shortage":
            shortage_col = cell.column
            break
    if shortage_col:
        for r in range(2, ws.max_row + 1):
            val = ws.cell(r, shortage_col).value
            try:
                if val and float(val) > 0:
                    for c in range(1, ws.max_column + 1):
                        ws.cell(r, c).fill = _RISK_FILL
            except Exception:
                pass

    # Автоширина
    widths = {}
    for r in ws.iter_rows(values_only=True):
        for idx, v in enumerate(r, start=1):
            w = len(str(v)) if v is not None else 0
            widths[idx] = max(widths.get(idx, 0), w)
    for idx, w in widths.items():
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = min(max(w + 2, 10), 60)


def _auto_width_template(ws):
    widths = {}
    for row in ws.iter_rows(values_only=True):
        for idx, value in enumerate(row, start=1):
            length = len(str(value)) if value is not None else 0
            widths[idx] = max(widths.get(idx, 0), length)
    for idx, width in widths.items():
        column = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[column].width = max(width + 2, 10)


def generate_input_template() -> io.BytesIO:
    wb = Workbook()

    input_headers = [
        "sku",
        "stock_ff",
        "stock_mp",
        "plan_sales_per_day",
        "safety_stock_ff",
        "safety_stock_mp",
    ]
    settings_headers = [
        "prod_lead_time_days",
        "lead_time_cn_msk",
        "lead_time_msk_mp",
        "moq_step_default",
        "safety_stock_ff_default",
        "safety_stock_mp_default",
    ]

    ws_input = wb.active
    ws_input.title = "Input"
    ws_input.append(input_headers)
    for cell in ws_input[1]:
        cell.font = _BOLD

    ws_input.append([
        "TEST_SKU",
        1000,
        800,
        12.5,
        "",  # override не задан
        "",  # override не задан
    ])
    _auto_width_template(ws_input)

    ws_settings = wb.create_sheet(SETTINGS_SHEET_NAME)
    ws_settings.append(settings_headers)
    for cell in ws_settings[1]:
        cell.font = _BOLD

    ws_settings.append([
        45,
        18,
        5,
        10,
        600,
        500,
    ])
    _auto_width_template(ws_settings)

    ws_transit = wb.create_sheet("InTransit")
    ws_transit.append(["sku", "qty", "eta_cn_msk"])
    for cell in ws_transit[1]:
        cell.font = _BOLD

    ws_transit.append([
        "TEST_SKU",
        120,
        "2025-11-01",
    ])
    _auto_width_template(ws_transit)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def build_output(xlsx_in: bytes, recs: List[Recommendation]) -> bytes:
    # Конвертируем в DataFrame и отсортируем колонки
    df_rec = pd.DataFrame([r.model_dump() for r in recs])
    if not df_rec.empty:
        df_rec = _order_columns(df_rec)

    in_buf = io.BytesIO(xlsx_in)
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as w:
        # Переносим исходные листы как есть (если читаются)
        try:
            xl = pd.ExcelFile(in_buf)
            for name in xl.sheet_names:
                pd.read_excel(xl, name).to_excel(w, sheet_name=name, index=False)
        except Exception:
            pass

        # Пишем Recommendations
        df_rec.to_excel(w, sheet_name="Recommendations", index=False)
        ws = w.book["Recommendations"]
        _apply_formats(ws)
        ws.freeze_panes = "A2"

    return out_buf.getvalue()

def process_excel(xlsx_bytes: bytes) -> bytes:
    items, trans = read_input(xlsx_bytes)
    recs = calculate(items, trans)
    return build_output(xlsx_bytes, recs)
