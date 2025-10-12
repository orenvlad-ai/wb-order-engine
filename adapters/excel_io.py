import io
from datetime import date, datetime
from typing import Tuple, List

import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment

from pydantic import ValidationError

from engine.models import SkuInput, InTransitItem, Recommendation
from engine.calc import calculate

REQUIRED_INPUT_COLS = [
    "sku","stock_ff","stock_mp","plan_sales_per_day",
    "prod_lead_time_days","lead_time_cn_msk","lead_time_msk_mp",
    "safety_stock_mp","moq_step"
]
REQUIRED_INTRANSIT_COLS = ["sku","qty","eta_cn_msk"]

class BadTemplateError(Exception): ...

def _ensure_columns(df: pd.DataFrame, required: List[str], sheet: str):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise BadTemplateError(f"Лист '{sheet}': нет колонок {missing}. Проверь шаблон.")

def read_input(xlsx_bytes: bytes) -> Tuple[List[SkuInput], List[InTransitItem]]:
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    try:
        df_in = pd.read_excel(xl, "Input")
        df_tr = pd.read_excel(xl, "InTransit")
    except ValueError as e:
        raise BadTemplateError("Нет листов 'Input' и/или 'InTransit' в файле.") from e

    _ensure_columns(df_in, REQUIRED_INPUT_COLS, "Input")
    _ensure_columns(df_tr, REQUIRED_INTRANSIT_COLS, "InTransit")

    items: List[SkuInput] = []
    for r in df_in.to_dict("records"):
        try:
            items.append(SkuInput(
                sku=str(r["sku"]),
                stock_ff=int(r["stock_ff"]),
                stock_mp=int(r["stock_mp"]),
                plan_sales_per_day=float(r["plan_sales_per_day"]),
                prod_lead_time_days=int(r["prod_lead_time_days"]),
                lead_time_cn_msk=int(r["lead_time_cn_msk"]),
                lead_time_msk_mp=int(r["lead_time_msk_mp"]),
                safety_stock_mp=int(r["safety_stock_mp"]),
                moq_step=int(r["moq_step"]),
            ))
        except (ValueError, ValidationError) as e:
            raise BadTemplateError(f"Строка Input для sku={r.get('sku')} содержит неверные данные.") from e

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
