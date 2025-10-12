from datetime import date, timedelta
import math
from typing import Iterable, List
from .models import SkuInput, InTransitItem, Recommendation
from . import ALGO_VERSION

def _today() -> date:
    return date.today()

def _calc_H(x: SkuInput) -> int:
    return x.prod_lead_time_days + x.lead_time_cn_msk + x.lead_time_msk_mp

def _eta_to_mp(it: InTransitItem, lt_msk_mp: int) -> date:
    return it.eta_cn_msk + timedelta(days=lt_msk_mp)

def _inbound_within_H(sku: str, items: Iterable[InTransitItem], lt_msk_mp: int, H: int, today: date) -> int:
    cutoff = today + timedelta(days=H - lt_msk_mp)
    return sum(it.qty for it in items if it.sku == sku and it.eta_cn_msk <= cutoff)

def _order_qty(shortage: float, moq_step: int) -> int:
    if shortage <= 0:
        return 0
    return int(math.ceil(shortage / moq_step) * moq_step)

def calculate(inputs: List[SkuInput], in_transit: List[InTransitItem]) -> List[Recommendation]:
    t = _today()
    recs: List[Recommendation] = []
    for x in inputs:
        H = _calc_H(x)
        demand = x.plan_sales_per_day * H
        inbound = _inbound_within_H(x.sku, in_transit, x.lead_time_msk_mp, H, t)
        coverage = x.stock_ff + x.stock_mp + inbound
        target = demand + x.safety_stock_mp
        shortage = max(0.0, target - coverage)
        order = _order_qty(shortage, x.moq_step)
        comment = f"H={H}; спрос={demand:.0f}; в_пути={inbound}; покрытие={coverage}; цель={target:.0f}; нехватка={shortage:.0f}"
        recs.append(Recommendation(
            sku=x.sku, H_days=H, demand_H=demand, inbound=inbound, coverage=coverage,
            target=target, shortage=shortage, moq_step=x.moq_step, order_qty=order,
            comment=comment, algo_version=ALGO_VERSION
        ))
    return recs
