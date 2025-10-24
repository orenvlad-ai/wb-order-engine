from datetime import date
from typing import Optional

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
    stock_status: str
    algo_version: str = "v1.2a"
    # Справочные остатки
    eoh: Optional[float] = None                    # Ост. к прих. заказа (до РП), шт
    eop_first: Optional[float] = None              # Остаток после первой поставки, шт
    stock_before_1: Optional[float] = None         # Ост. до 1П
    stock_after_1: Optional[float] = None          # Ост. после 1П
    stock_before_2: Optional[float] = None         # Ост. до 2П
    stock_after_2: Optional[float] = None          # Ост. после 2П
    stock_before_3: Optional[float] = None         # Ост. до 3П
    stock_after_3: Optional[float] = None          # Ост. после 3П
    stock_before_po: Optional[float] = None        # Ост. до РП
    stock_after_po: Optional[float] = None         # Ост. после РП
