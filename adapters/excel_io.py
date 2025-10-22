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
REQUIRED_INTRANSIT_COLS = ["sku", "qty", "eta_cn_msk"]
SETTINGS_SHEET_NAME = "Настройки заказа"

INPUT_SHEET_NAMES = ("Input", "Ввод")
INTRANSIT_SHEET_NAMES = ("InTransit", "Товары в пути")

# Отображения колонок листа Recommendations
RECOMMENDATION_COLUMN_ALIASES = {
    "sku": "Артикул",
    "order_qty": "Рекомендуемый заказ, шт",
    "shortage": "Нехватка, шт",
    "target": "Цель, шт",
    "coverage": "Покрытие, шт",
    "inbound": "В пути, шт",
    "demand_H": "Спрос за горизонт, шт",
    "H_days": "Горизонт прогноза, дней",
    "moq_step": "Кратность заказа (MOQ)",
    "stock_status": "Статус запаса",
    "reduce_plan_to": "Рекоменд. план, шт/день",
    "comment": "Комментарий",
    "algo_version": "Версия алгоритма",
}

RECOMMENDATION_DISPLAY_TO_INTERNAL = {
    v: k for k, v in RECOMMENDATION_COLUMN_ALIASES.items()
}

# Сопоставление русских заголовков с внутренними ключами
INPUT_COLUMN_ALIASES = {
    "Артикул": "sku",
    "Остаток ФФ": "stock_ff",
    "Остаток МП": "stock_mp",
    "План, шт/день": "plan_sales_per_day",
    "Несниж. остаток ФФ": "safety_stock_ff",
    "Несниж. остаток МП": "safety_stock_mp",
}

INTRANSIT_COLUMN_ALIASES = {
    "Артикул": "sku",
    "Кол-во": "qty",
    "ETA на ФФ": "eta_cn_msk",
}

SETTINGS_COLUMN_ALIASES = {
    "Произв., дней": "prod_lead_time_days",
    "Китай→МСК, дней": "lead_time_cn_msk",
    "МСК→МП, дней": "lead_time_msk_mp",
    "Кратность (MOQ)": "moq_step_default",
    "Порог несниж. МП при OOS, %": "oos_safety_mp_pct",
    "Дефолт. несниж. ФФ": "safety_stock_ff_default",
    "Дефолт. несниж. МП": "safety_stock_mp_default",
}

# Отображение внутренних имён обратно в русские заголовки для сообщений об ошибках
INPUT_COLUMN_DISPLAY = {
    "sku": "Артикул",
    "stock_ff": "Остаток ФФ",
    "stock_mp": "Остаток МП",
    "plan_sales_per_day": "План, шт/день",
    "safety_stock_ff": "Несниж. остаток ФФ",
    "safety_stock_mp": "Несниж. остаток МП",
}

INTRANSIT_COLUMN_DISPLAY = {
    "sku": "Артикул",
    "qty": "Кол-во",
    "eta_cn_msk": "ETA на ФФ",
}

SETTINGS_COLUMN_DISPLAY = {
    "prod_lead_time_days": "Произв., дней",
    "lead_time_cn_msk": "Китай→МСК, дней",
    "lead_time_msk_mp": "МСК→МП, дней",
    "moq_step_default": "Кратность (MOQ)",
    "oos_safety_mp_pct": "Порог несниж. МП при OOS, %",
    "safety_stock_ff_default": "Дефолт. несниж. ФФ",
    "safety_stock_mp_default": "Дефолт. несниж. МП",
}
REQUIRED_SETTINGS_COLS = [
    "prod_lead_time_days",
    "lead_time_cn_msk",
    "lead_time_msk_mp",
    "moq_step_default",
    "oos_safety_mp_pct",
]
OPTIONAL_SETTINGS_COLS = [
    "safety_stock_ff_default",
    "safety_stock_mp_default",
]

class BadTemplateError(Exception): ...


def _ensure_columns(
    df: pd.DataFrame,
    required: List[str],
    sheet: str,
    display_names: Optional[Dict[str, str]] = None,
):
    missing = [c for c in required if c not in df.columns]
    if missing:
        readable = [display_names.get(c, c) for c in missing] if display_names else missing
        raise BadTemplateError(
            f"На листе '{sheet}' отсутствуют обязательные колонки: {', '.join(readable)}."
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


def _read_settings(df_settings: pd.DataFrame) -> Dict[str, Any]:
    # Считываем первую заполненную строку с общими параметрами заказа
    _ensure_columns(
        df_settings,
        REQUIRED_SETTINGS_COLS,
        SETTINGS_SHEET_NAME,
        SETTINGS_COLUMN_DISPLAY,
    )
    df_clean = df_settings.dropna(how="all")
    if df_clean.empty:
        raise BadTemplateError(
            "Лист 'Настройки заказа' пуст. Заполни строку с параметрами по умолчанию."
        )

    row = df_clean.iloc[0].to_dict()
    settings: Dict[str, Any] = {}
    for key in REQUIRED_SETTINGS_COLS:
        value = row.get(key)
        if key == "oos_safety_mp_pct":
            value = 5 if _is_blank(value) else value
            settings[key] = _parse_float(
                value,
                sheet=SETTINGS_SHEET_NAME,
                column=SETTINGS_COLUMN_DISPLAY.get(key, key),
                sku="*",
            )
        else:
            settings[key] = _parse_int(
                value,
                sheet=SETTINGS_SHEET_NAME,
                column=SETTINGS_COLUMN_DISPLAY.get(key, key),
            )
    for key in OPTIONAL_SETTINGS_COLS:
        if key in df_settings.columns:
            value = row.get(key)
            settings[key] = _parse_int(
                value,
                sheet=SETTINGS_SHEET_NAME,
                column=SETTINGS_COLUMN_DISPLAY.get(key, key),
            )
        else:
            settings[key] = 0
    return settings

def read_input(xlsx_bytes: bytes) -> Tuple[List[SkuInput], List[InTransitItem]]:
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    input_sheet_name = next((name for name in INPUT_SHEET_NAMES if name in xl.sheet_names), None)
    if input_sheet_name is None:
        raise BadTemplateError("В файле нет листа 'Input' или 'Ввод'.")

    df_in = pd.read_excel(xl, input_sheet_name)
    df_in.rename(columns=INPUT_COLUMN_ALIASES, inplace=True)
    df_in = df_in.where(pd.notna(df_in), None)

    try:
        df_settings = pd.read_excel(xl, SETTINGS_SHEET_NAME)
    except ValueError as e:
        raise BadTemplateError("В файле нет листа 'Настройки заказа'.") from e

    df_settings.rename(columns=SETTINGS_COLUMN_ALIASES, inplace=True)
    df_settings = df_settings.where(pd.notna(df_settings), None)

    _ensure_columns(
        df_in,
        REQUIRED_INPUT_COLS,
        input_sheet_name,
        INPUT_COLUMN_DISPLAY,
    )
    for col in OPTIONAL_INPUT_COLS:
        if col not in df_in.columns:
            df_in[col] = None

    settings = _read_settings(df_settings)

    transit_sheet_name = next(
        (name for name in INTRANSIT_SHEET_NAMES if name in xl.sheet_names),
        None,
    )
    try:
        if transit_sheet_name is None:
            raise ValueError
        df_tr = pd.read_excel(xl, transit_sheet_name)
        df_tr.rename(columns=INTRANSIT_COLUMN_ALIASES, inplace=True)
        df_tr = df_tr.where(pd.notna(df_tr), None)
        _ensure_columns(
            df_tr,
            REQUIRED_INTRANSIT_COLS,
            transit_sheet_name,
            INTRANSIT_COLUMN_DISPLAY,
        )
    except ValueError:
        df_tr = pd.DataFrame(columns=REQUIRED_INTRANSIT_COLS)
    except BadTemplateError:
        raise

    items: List[SkuInput] = []
    moq_step_default = settings["moq_step_default"]
    safety_stock_ff_default = settings["safety_stock_ff_default"]
    safety_stock_mp_default = settings["safety_stock_mp_default"]
    oos_safety_mp_pct = float(settings.get("oos_safety_mp_pct", 5))
    prod_lead_time_days = settings["prod_lead_time_days"]
    lead_time_cn_msk = settings["lead_time_cn_msk"]
    lead_time_msk_mp = settings["lead_time_msk_mp"]

    for r in df_in.to_dict("records"):
        if all(_is_blank(v) for v in r.values()):
            continue
        sku = str(r.get("sku") or "").strip()
        if not sku:
            raise BadTemplateError(
                f"На листе '{input_sheet_name}' есть строка без SKU. Удали пустые строки."
            )
        try:
            stock_ff = _parse_int(
                r.get("stock_ff"),
                sheet=input_sheet_name,
                column=INPUT_COLUMN_DISPLAY["stock_ff"],
                sku=sku,
            )
            stock_mp = _parse_int(
                r.get("stock_mp"),
                sheet=input_sheet_name,
                column=INPUT_COLUMN_DISPLAY["stock_mp"],
                sku=sku,
            )
            plan_sales_per_day = _parse_float(
                r.get("plan_sales_per_day"),
                sheet=input_sheet_name,
                column=INPUT_COLUMN_DISPLAY["plan_sales_per_day"],
                sku=sku,
            )

            raw_mp = r.get("safety_stock_mp")
            if _is_blank(raw_mp):
                safety_stock_mp = safety_stock_mp_default  # пусто → берём дефолт из настроек
            else:
                safety_stock_mp = _parse_int(
                    raw_mp,
                    sheet=input_sheet_name,
                    column=INPUT_COLUMN_DISPLAY["safety_stock_mp"],
                    sku=sku,
                )

            raw_ff = r.get("safety_stock_ff")
            if _is_blank(raw_ff):
                safety_stock_ff = safety_stock_ff_default  # пусто → берём дефолт из настроек
            else:
                safety_stock_ff = _parse_int(
                    raw_ff,
                    sheet=input_sheet_name,
                    column=INPUT_COLUMN_DISPLAY["safety_stock_ff"],
                    sku=sku,
                )

            items.append(SkuInput(
                sku=sku,
                stock_ff=stock_ff,
                stock_mp=stock_mp,
                plan_sales_per_day=plan_sales_per_day,
                prod_lead_time_days=prod_lead_time_days,
                lead_time_cn_msk=lead_time_cn_msk,
                lead_time_msk_mp=lead_time_msk_mp,
                oos_safety_mp_pct=oos_safety_mp_pct,
                safety_stock_mp=safety_stock_mp,
                safety_stock_ff=safety_stock_ff,
                moq_step=moq_step_default,
            ))
        except (ValueError, ValidationError) as e:
            raise BadTemplateError(
                f"На листе '{input_sheet_name}' строка для SKU '{sku}' содержит неверные данные."
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
            raise BadTemplateError(
                f"На листе '{transit_sheet_name or INTRANSIT_SHEET_NAMES[0]}' строка для SKU "
                f"'{r.get('sku')}' имеет неверные количество или дату."
            ) from e

    return items, trans

# ---------- Форматирование Recommendations ----------

_HEADER_FILL = PatternFill(start_color="FFEFEFEF", end_color="FFEFEFEF", fill_type="solid")
_RISK_FILL   = PatternFill(start_color="FFFFE5E5", end_color="FFFFE5E5", fill_type="solid")
_BOLD = Font(bold=True)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
_THIN = Side(border_style="thin", color="FFBFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# Порядок колонок (берём те, что реально есть в данных)
_ORDER = [
    "sku", "order_qty", "stock_status", "reduce_plan_to", "comment",
    "shortage", "target", "coverage", "inbound",
    "demand_H", "H_days", "moq_step", "algo_version"
]

def _order_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in _ORDER if c in df.columns] + [c for c in df.columns if c not in _ORDER]
    return df[cols]

def _find_col_idx_by_internal(ws, internal_key: str) -> int | None:
    header_rows = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    if not header_rows:
        return None
    for idx, value in enumerate(header_rows[0], start=1):
        if value == internal_key:
            return idx
        if RECOMMENDATION_DISPLAY_TO_INTERNAL.get(value) == internal_key:
            return idx
    return None


def _apply_formats_localized(ws):
    idx_order = _find_col_idx_by_internal(ws, "order_qty")
    idx_status = _find_col_idx_by_internal(ws, "stock_status")
    idx_short = _find_col_idx_by_internal(ws, "shortage")
    idx_cov = _find_col_idx_by_internal(ws, "coverage")
    _apply_formats(
        ws,
        idx_order=idx_order,
        idx_status=idx_status,
        idx_short=idx_short,
        idx_cov=idx_cov,
    )


def _apply_formats(
    ws,
    *,
    idx_order: int | None = None,
    idx_status: int | None = None,
    idx_short: int | None = None,
    idx_cov: int | None = None,
):
    header_internal: Dict[int, str] = {}

    # Шапка
    for idx, cell in enumerate(ws[1], start=1):
        internal_name = RECOMMENDATION_DISPLAY_TO_INTERNAL.get(cell.value, cell.value)
        header_internal[idx] = internal_name
        cell.font = _BOLD
        cell.alignment = _CENTER  # центр по горизонтали и вертикали
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
        # краткие подсказки
        if internal_name == "order_qty":
            cell.comment = Comment("Рекомендованный заказ с кратностью MOQ", "WB Engine")
        if internal_name == "shortage":
            cell.comment = Comment("Нехватка к цели (demand_H + safety_stock)", "WB Engine")

    for col_idx, name in (
        (idx_order, "order_qty"),
        (idx_status, "stock_status"),
        (idx_short, "shortage"),
        (idx_cov, "coverage"),
    ):
        if col_idx and header_internal.get(col_idx) != name:
            header_internal[col_idx] = name

    # Форматы чисел и выравнивание данных
    int_like = {
        "H_days",
        "inbound",
        "coverage",
        "target",
        "shortage",
        "moq_step",
        "order_qty",
        "demand_H",
        "reduce_plan_to",
    }
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            header_value = header_internal.get(cell.column)
            cell.border = _BORDER
            if header_value in int_like:
                cell.number_format = "0"
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=False)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

    # Подсветка: красный фон только при статусе ⚠️
    if idx_status:
        for r in range(2, ws.max_row + 1):
            status_val = str(ws.cell(r, idx_status).value or "")
            if "⚠️" in status_val:
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = _RISK_FILL

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
        "Артикул",
        "Остаток ФФ",
        "Остаток МП",
        "План, шт/день",
        "Несниж. остаток ФФ",
        "Несниж. остаток МП",
    ]
    settings_headers = [
        "Произв., дней",
        "Китай→МСК, дней",
        "МСК→МП, дней",
        "Кратность (MOQ)",
        "Порог несниж. МП при OOS, %",
    ]

    ws_input = wb.active
    ws_input.title = "Ввод"
    ws_input.append(input_headers)
    for cell in ws_input[1]:
        cell.font = _BOLD

    ws_input.append([
        "SKU123",
        900,
        650,
        14.5,
        "",
        "",
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
        5,
    ])
    _auto_width_template(ws_settings)

    ws_transit = wb.create_sheet("Товары в пути")
    ws_transit.append(["Артикул", "Кол-во", "ETA на ФФ"])
    for cell in ws_transit[1]:
        cell.font = _BOLD

    ws_transit.append([
        "SKU123",
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
        # 1) Пишем Recommendations ПЕРВЫМ листом
        df_out = df_rec.rename(columns=RECOMMENDATION_COLUMN_ALIASES)
        df_out.to_excel(w, sheet_name="Recommendations", index=False)
        ws = w.book["Recommendations"]
        _apply_formats_localized(ws)
        ws.freeze_panes = "A2"

        # 2) Пишем скрытый лист Log с техполями
        log_cols = [
            "sku",
            "H_days",
            "demand_H",
            "inbound",
            "coverage",
            "target",
            "shortage",
            "moq_step",
            "order_qty",
            "stock_status",
            "reduce_plan_to",
            "algo_version",
        ]
        log_df = df_rec.reindex(columns=log_cols)
        if log_df.shape[1]:
            log_df.to_excel(w, sheet_name="Log", index=False)
            ws_log = w.book["Log"]
            ws_log.sheet_state = "hidden"

        # 3) Затем переносим прочие исходные листы
        try:
            xl = pd.ExcelFile(in_buf)
            for name in xl.sheet_names:
                if name in ("Recommendations", "Log"):
                    continue
                pd.read_excel(xl, name).to_excel(w, sheet_name=name, index=False)
        except Exception:
            pass

    return out_buf.getvalue()

def process_excel(xlsx_bytes: bytes) -> bytes:
    items, trans = read_input(xlsx_bytes)
    recs = calculate(items, trans)
    return build_output(xlsx_bytes, recs)
