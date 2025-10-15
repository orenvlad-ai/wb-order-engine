# app/main.py
from datetime import date, datetime
from io import BytesIO
from typing import List

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook

from engine.calc import calculate
from engine.models import SkuInput, InTransitItem
from engine.config import ALGO_VERSION
import uvicorn


app = FastAPI(title="WB Order Engine")

# Шаблоны и статика
templates = Jinja2Templates(directory="app/templates")
app.mount("/static", StaticFiles(directory="app/static"), name="static")


# Форма ввода (GET)
@app.get("/", response_class=HTMLResponse)
async def input_form(request: Request):
    return templates.TemplateResponse("input_form.html", {"request": request})


# Обработчик кнопки "Рассчитать и скачать Excel" (POST)
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

    # Повторяющиеся поля из формы (динамические строки "Товар в пути")
    in_transit_qty: List[int] = Form(default=[]),
    in_transit_eta: List[str] = Form(default=[]),
):
    # 1) Сбор входных данных
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

    # 2) Парсим партии в пути (если пользователь добавил строки)
    in_transit: List[InTransitItem] = []
    for q, eta_str in zip(in_transit_qty or [], in_transit_eta or []):
        if not eta_str:
            continue
        eta = datetime.strptime(eta_str, "%Y-%m-%d").date()
        in_transit.append(InTransitItem(sku=item.sku, qty=int(q), eta_cn_msk=eta))

    # 3) Расчёт рекомендаций
    recs = calculate([item], in_transit=in_transit)

    # 4) Сборка Excel: лист Recommendations
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

    # 5) Отдаём файл пользователю
    buff = BytesIO()
    wb.save(buff)
    buff.seek(0)

    filename = f"Planner_Recommendations_{date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buff,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)
