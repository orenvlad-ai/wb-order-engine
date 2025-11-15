# app/main.py
from datetime import datetime
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from engine.calc import calculate
from engine.config import ALGO_VERSION
from adapters.excel_io import BadTemplateError, build_output, generate_input_template, read_input
import uvicorn
import logging

LAST_RESULTS_DIR = Path("last_results")
LAST_RESULTS_DIR.mkdir(parents=True, exist_ok=True)


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

        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"Planner_Recommendations_{now}.xlsx"
        result_path = LAST_RESULTS_DIR / fname
        with open(result_path, "wb") as f_out:
            f_out.write(out_bytes)

        try:
            files = sorted(
                [p for p in LAST_RESULTS_DIR.glob("*.xlsx") if p.is_file()],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            for old in files[5:]:
                try:
                    old.unlink()
                except OSError:
                    continue
        except Exception:
            pass

        buffer = BytesIO(out_bytes)
        return StreamingResponse(
            buffer,
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


@app.get("/last_results")
async def list_last_results():
    try:
        files = sorted(
            [p for p in LAST_RESULTS_DIR.glob("*.xlsx") if p.is_file()],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )[:5]
    except Exception:
        files = []

    result = []
    for p in files:
        try:
            stat = p.stat()
            result.append(
                {
                    "name": p.name,
                    "size": stat.st_size,
                    "mtime": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                }
            )
        except Exception:
            continue
    return JSONResponse(result)


@app.get("/last_results/{filename}")
async def download_last_result(filename: str):
    path = LAST_RESULTS_DIR / filename
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Файл не найден")

    def _iterfile():
        with open(path, "rb") as f:
            while True:
                chunk = f.read(8192)
                if not chunk:
                    break
                yield chunk

    return StreamingResponse(
        _iterfile(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ---------------------- Локальный запуск ----------------------
if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)
