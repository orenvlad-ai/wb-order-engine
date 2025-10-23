from datetime import date
from typing import Optional, Union

from pydantic import BaseModel, Field


class SkuInput(BaseModel):
    sku: str
    stock_ff: int = Field(ge=0)
    stock_mp: int = Field(ge=0)
    plan_sales_per_day: float = Field(ge=0)
    prod_lead_time_days: int = Field(ge=0)
    lead_time_cn_msk: int = Field(ge=0)
    lead_time_msk_mp: int = Field(ge=0)
    oos_safety_mp_pct: float = Field(ge=0, le=100, default=5.0)
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
    stock_status: Optional[str] = None
    reduce_plan_to: Optional[Union[float, str]] = None
    reduce_plan_to_after: Optional[Union[float, str]] = None
    comment: str
    algo_version: str = "v1.2a"
    eoh: Optional[float] = None  # Остаток на конец горизонта
    eop_first: Optional[float] = None  # Остаток к первой интранзит-поставке в пределах горизонта
    debug_r1_smooth: Optional[float] = None
    debug_r2_smooth: Optional[float] = None
    debug_d1: Optional[float] = None
    debug_d2: Optional[float] = None
    debug_demand_first: Optional[float] = None
    debug_demand_after: Optional[float] = None
    debug_eoh_before: Optional[float] = None
    debug_eoh_after: Optional[float] = None
