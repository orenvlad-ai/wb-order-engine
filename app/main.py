from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List
from datetime import datetime
from engine.models import SkuInput, InTransitItem, Recommendation
from engine.calc import calculate
from adapters.excel_io import process_excel, BadTemplateError
import io

app = FastAPI(title="WB Order Engine", version="0.2.0")

class Payload(BaseModel):
    items: List[SkuInput]
    in_transit: List[InTransitItem] = []

@app.get("/")
def root():
    return {"status": "ok", "message": "WB Order Engine is alive 🚀"}

@app.post("/api/recommendations/json", response_model=List[Recommendation])
def rec_json(payload: Payload):
    return calculate(payload.items, payload.in_transit)

@app.post("/api/recommendations/excel")
async def rec_excel(file: UploadFile = File(...)):
    try:
        content = await file.read()
        result = process_excel(content)
    except BadTemplateError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception:
        raise HTTPException(status_code=500, detail="Ошибка обработки файла")

    stamp = datetime.now().date().isoformat()
    filename = f"Planner_Recommendations_{stamp}.xlsx"
    return StreamingResponse(
        io.BytesIO(result),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
