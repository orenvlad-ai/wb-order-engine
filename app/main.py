# app/main.py
from datetime import date
from io import BytesIO

from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from engine.calc import calculate
from engine.config import ALGO_VERSION
from adapters.excel_io import BadTemplateError, build_output, generate_input_template, read_input
import uvicorn
import logging


app = FastAPI(title="WB Order Engine")

# --- Шаблоны и статика ---
templates = Jinja2Templates(directory="app/templates")
app.mount("/static", StaticFiles(directory="app/static"), name="static")

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
    buffer = generate_input_template()
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Input_Template.xlsx"'},
    )


@app.get("/download_input_template")
async def download_input_template():
    buffer = generate_input_template()
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Input_Template.xlsx"'},
    )


# ---------------------- Загрузка Excel -> Excel ----------------------
@app.post("/upload_excel")
async def upload_excel(file: UploadFile = File(...)):
    try:
        if not file.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail="Ожидается .xlsx файл.")

        content = await file.read()
        items, in_transit = read_input(content)
        recs = calculate(items, in_transit)
        out_bytes = build_output(content, recs)

        fname = f"Planner_Recommendations_{date.today().isoformat()}.xlsx"
        return StreamingResponse(
            BytesIO(out_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'}
        )
    except HTTPException as exc:
        return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})
    except BadTemplateError as exc:
        return JSONResponse(status_code=400, content={"detail": str(exc)})
    except Exception:
        logging.exception("Unexpected error while processing Excel upload")
        return JSONResponse(
            status_code=500,
            content={"detail": "Внутренняя ошибка при обработке Excel"}
        )


# ---------------------- Локальный запуск ----------------------
if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)
