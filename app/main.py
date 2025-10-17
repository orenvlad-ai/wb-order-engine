# app/main.py
from datetime import date, datetime
from io import BytesIO
from typing import List, Tuple
from pathlib import Path

from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from openpyxl import Workbook, load_workbook

from engine.calc import calculate
from engine.models import SkuInput, InTransitItem
from engine.config import ALGO_VERSION
import uvicorn
import logging


app = FastAPI(title="WB Order Engine")

# --- Шаблоны и статика ---
templates = Jinja2Templates(directory="app/templates")
app.mount("/static", StaticFiles(directory="app/static"), name="static")

# --- Путь к Excel-шаблону ---
TEMPLATE_FILE = Path(__file__).parent / "static" / "templates" / "Input_Template_Items_InTransit.xlsx"

# --- UI: форма ввода ---
@app.get("/", response_class=HTMLResponse)
async def input_form(request: Request):
    return templates.TemplateResponse(
        "input_form.html",
        {"request": request, "algo_version": ALGO_VERSION},
    )


@app.get("/health")
async def health_check():
    return {"status": "ok", "algo_version": ALGO_VERSION}


# ---------------------- Скачивание шаблона ----------------------
@app.get("/download_template")
async def download_template():
    if not TEMPLATE_FILE.exists():
        raise HTTPException(status_code=404, detail="Template not found")
    return FileResponse(
        path=str(TEMPLATE_FILE),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Input_Template_Items_InTransit.xlsx",
    )


# ---------------------- Вспомогательные функции ----------------------
REQUIRED_ITEMS_COLS = [
    "sku", "stock_ff", "stock_mp", "plan_sales_per_day", "prod_lead_time_days",
    "lead_time_cn_msk", "lead_time_msk_mp", "safety_stock_mp", "moq_step"
]
INTRANSIT_COLS = ["sku", "qty", "eta_cn_msk"]


def _coerce_date(v) -> date:
    if isinstance(v, date):
        return v
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, str):
        return datetime.strptime(v.strip(), "%Y-%m-%d").date()
    raise ValueError(f"Некорректная дата: {v!r}")


def _read_items(ws) -> List[SkuInput]:
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    missing = [c for c in REQUIRED_ITEMS_COLS if c not in headers]
    if missing:
        raise HTTPException(status_code=400, detail=f"Лист 'Items': нет колонок: {', '.join(missing)}")

    idx = {name: headers.index(name) for name in REQUIRED_ITEMS_COLS}
    items: List[SkuInput] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue

        def get(name):
            j = idx[name]
            return row[j]

        try:
            items.append(SkuInput(
                sku=str(get("sku")).strip(),
                stock_ff=int(get("stock_ff") or 0),
                stock_mp=int(get("stock_mp") or 0),
                plan_sales_per_day=float(get("plan_sales_per_day") or 0),
                prod_lead_time_days=int(get("prod_lead_time_days") or 0),
                lead_time_cn_msk=int(get("lead_time_cn_msk") or 0),
                lead_time_msk_mp=int(get("lead_time_msk_mp") or 0),
                safety_stock_mp=int(get("safety_stock_mp") or 0),
                moq_step=int(get("moq_step") or 1),
            ))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Лист 'Items': ошибка парсинга строки {row}: {e}")
    if not items:
        raise HTTPException(status_code=400, detail="Лист 'Items' пуст.")
    return items


def _read_intransit(ws) -> List[InTransitItem]:
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    missing = [c for c in INTRANSIT_COLS if c not in headers]
    if missing:
        raise HTTPException(status_code=400, detail=f"Лист 'InTransit': нет колонок: {', '.join(missing)}")

    idx = {name: headers.index(name) for name in INTRANSIT_COLS}
    rows: List[InTransitItem] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue

        def get(name):
            j = idx[name]
            return row[j]

        try:
            rows.append(InTransitItem(
                sku=str(get("sku")).strip(),
                qty=int(get("qty") or 0),
                eta_cn_msk=_coerce_date(get("eta_cn_msk")),
            ))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Лист 'InTransit': ошибка парсинга строки {row}: {e}")
    return rows


def _parse_input_excel(content: bytes) -> Tuple[List[SkuInput], List[InTransitItem]]:
    try:
        wb = load_workbook(filename=BytesIO(content), data_only=True)
    except Exception:
        raise HTTPException(status_code=400, detail="Не удалось открыть Excel. Убедись, что это .xlsx файл.")

    if "Items" not in wb.sheetnames:
        raise HTTPException(status_code=400, detail="В файле отсутствует лист 'Items'.")

    items = _read_items(wb["Items"])
    in_transit: List[InTransitItem] = []
    if "InTransit" in wb.sheetnames:
        in_transit = _read_intransit(wb["InTransit"])
    return items, in_transit


def _excel_from_recs(recs):
    wb = Workbook()
    ws = wb.active
    ws.title = "Recommendations"

    headers = [
        "sku", "H_days", "demand_H", "inbound", "coverage",
        "target", "shortage", "moq_step", "order_qty",
        "reduce_plan_to", "comment", "algo_version"
    ]
    ws.append(headers)

    for r in recs:
        ws.append([
            r.sku, r.H_days, r.demand_H, r.inbound, r.coverage,
            r.target, r.shortage, r.moq_step, r.order_qty,
            getattr(r, "reduce_plan_to", None),
            r.comment, r.algo_version
        ])

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, max_len + 2), 40)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------- Ручной ввод -> Excel ----------------------
@app.post("/calc_excel")
async def calc_excel(
    sku: str = Form(...),
    stock_ff: int = Form(...),
    stock_mp: int = Form(...),
    plan_sales_per_day: float = Form(...),
    prod_lead_time_days: int = Form(...),
    lead_time_cn_msk: int = Form(...),
    lead_time_msk_mp: int = Form(...),
    safety_stock_mp: int = Form(...),
    moq_step: int = Form(...),

    in_transit_qty: List[int] = Form(default=[]),
    in_transit_eta: List[str] = Form(default=[]),
):
    item = SkuInput(
        sku=sku.strip(),
        stock_ff=stock_ff,
        stock_mp=stock_mp,
        plan_sales_per_day=plan_sales_per_day,
        prod_lead_time_days=prod_lead_time_days,
        lead_time_cn_msk=lead_time_cn_msk,
        lead_time_msk_mp=lead_time_msk_mp,
        safety_stock_mp=safety_stock_mp,
        moq_step=moq_step,
    )

    in_transit: List[InTransitItem] = []
    for q, eta_str in zip(in_transit_qty or [], in_transit_eta or []):
        if not eta_str:
            continue
        eta = datetime.strptime(eta_str, "%Y-%m-%d").date()
        in_transit.append(InTransitItem(sku=item.sku, qty=int(q), eta_cn_msk=eta))

    recs = calculate([item], in_transit=in_transit)
    buff = _excel_from_recs(recs)

    filename = f"Planner_Recommendations_{date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buff,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ---------------------- Загрузка Excel -> Excel ----------------------
@app.post("/upload_excel")
async def upload_excel(file: UploadFile = File(...)):
    try:
        if not file.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail="Ожидается .xlsx файл.")

        content = await file.read()
        items, in_transit = _parse_input_excel(content)
        recs = calculate(items, in_transit)
        buff = _excel_from_recs(recs)

        fname = f"Planner_Recommendations_{date.today().isoformat()}.xlsx"
        return StreamingResponse(
            buff,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'}
        )
    except HTTPException as exc:
        return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})
    except Exception:
        logging.exception("Unexpected error while processing Excel upload")
        return JSONResponse(
            status_code=500,
            content={"detail": "Внутренняя ошибка при обработке Excel"}
        )


# ---------------------- Локальный запуск ----------------------
if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)
