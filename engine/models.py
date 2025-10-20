from pydantic import BaseModel, Field
from datetime import date

class SkuInput(BaseModel):
    sku: str
    stock_ff: int = Field(ge=0)
    stock_mp: int = Field(ge=0)
    plan_sales_per_day: float = Field(ge=0)
    prod_lead_time_days: int = Field(ge=0)
    lead_time_cn_msk: int = Field(ge=0)
    lead_time_msk_mp: int = Field(ge=0)
    safety_stock_mp: int = Field(ge=0)
    safety_stock_ff: int = Field(ge=0)
    moq_step: int = Field(ge=1)

class InTransitItem(BaseModel):
    sku: str
    qty: int = Field(ge=0)
    eta_cn_msk: date

class Recommendation(BaseModel):
    sku: str
    H_days: int
    demand_H: float
    inbound: float
    coverage: float
    target: float
    shortage: float
    moq_step: int
    order_qty: int
    comment: str
    algo_version: str = "v1.2a"
