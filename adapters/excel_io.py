import io
from datetime import date
import pandas as pd
from pydantic import ValidationError
from typing import Tuple, List
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

def build_output(xlsx_in: bytes, recs: List[Recommendation]) -> bytes:
    # Берём исходные листы и добавляем Recommendations
    in_buf = io.BytesIO(xlsx_in)
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as w:
        # переносим существующие листы как есть
        try:
            xl = pd.ExcelFile(in_buf)
            for name in xl.sheet_names:
                pd.read_excel(xl, name).to_excel(w, sheet_name=name, index=False)
        except Exception:
            pass  # если не смогли прочитать — просто создадим Recommendations
        # Recommendations
        df_rec = pd.DataFrame([r.model_dump() for r in recs])
        df_rec.to_excel(w, sheet_name="Recommendations", index=False)
    return out_buf.getvalue()

def process_excel(xlsx_bytes: bytes) -> bytes:
    items, trans = read_input(xlsx_bytes)
    recs = calculate(items, trans)
    return build_output(xlsx_bytes, recs)
