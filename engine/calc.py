from datetime import date, timedelta
import math
from typing import Iterable, List, Optional, Tuple

from .models import SkuInput, InTransitItem, Recommendation
from .config import ALGO_VERSION

def _today() -> date:
    return date.today()

def _calc_H(x: SkuInput) -> int:
    return x.prod_lead_time_days + x.lead_time_cn_msk + x.lead_time_msk_mp

def _eta_to_mp(it: InTransitItem, lt_msk_mp: int) -> date:
    return it.eta_cn_msk + timedelta(days=lt_msk_mp)

def _inbound_within_H(
    sku: str,
    items: Iterable[InTransitItem],
    lt_msk_mp: int,
    H: int,
    today: date,
) -> Tuple[int, Optional[date]]:
    horizon_mp = today + timedelta(days=H)
    inbound = 0
    next_eta_mp: Optional[date] = None
    for it in items:
        if it.sku != sku:
            continue
        eta_mp = _eta_to_mp(it, lt_msk_mp)
        if eta_mp < today:
            continue
        if eta_mp <= horizon_mp:
            inbound += it.qty
        if next_eta_mp is None or eta_mp < next_eta_mp:
            next_eta_mp = eta_mp
    return inbound, next_eta_mp

def _order_qty(shortage: float, moq_step: int) -> int:
    if shortage <= 0:
        return 0
    return int(math.ceil(shortage / moq_step) * moq_step)

def calculate(inputs: List[SkuInput], in_transit: List[InTransitItem]) -> List[Recommendation]:
    t = _today()
    recs: List[Recommendation] = []
    for x in inputs:
        H = _calc_H(x)
        inbound, next_eta_mp = _inbound_within_H(
            x.sku, in_transit, x.lead_time_msk_mp, H, t
        )
        coverage = x.stock_ff + x.stock_mp + inbound

        # Определяем дни до ближайшей поставки на МП.
        # Если поставки нет, считаем, что доживать нужно ВЕСЬ горизонт (H).
        if next_eta_mp is None:
            days_until_next_inbound: float = H
        else:
            days_until_next_inbound = max((next_eta_mp - t).days, 0)

        on_hand = x.stock_ff + x.stock_mp
        oos_threshold = (x.oos_safety_mp_pct / 100.0) * x.safety_stock_mp
        usable = max(0.0, on_hand - oos_threshold)
        coverage_days_on_hand = (
            float("inf")
            if x.plan_sales_per_day <= 0
            else usable / x.plan_sales_per_day
        )

        if coverage_days_on_hand < days_until_next_inbound:
            stock_status = "⚠️ Не хватает до поставки"
            denom = max(days_until_next_inbound, 1)
            max_daily = usable / denom
            reduce_plan_to = float(
                max(
                    0.0,
                    math.floor(
                        min(
                            x.plan_sales_per_day,
                            max_daily,
                        )
                    ),
                )
            )
        else:
            stock_status = "✅ Запаса хватает до поставки"
            reduce_plan_to = None

        if stock_status.startswith("⚠️"):
            effective_plan_before = (
                reduce_plan_to if reduce_plan_to is not None else x.plan_sales_per_day
            )
            days_until = days_until_next_inbound
            demand_H = (
                effective_plan_before * days_until
                + x.plan_sales_per_day * max(0, H - days_until)
            )
        else:
            demand_H = x.plan_sales_per_day * H

        # Покрытие на горизонте (учитывает только интранзиты, успевающие до H)
        target = demand_H + x.safety_stock_mp + x.safety_stock_ff
        shortage = max(0.0, target - coverage)
        order_qty = _order_qty(shortage, x.moq_step)

        # EOH как остаток НА МОМЕНТ ПРИБЫТИЯ РАСЧЁТНОГО ЗАКАЗА (через H дней):
        # (Новая поставка order_qty в это значение НЕ включается.)
        #   eoh = on_hand + inbound_≤H − demand_до(H)
        # где demand_до(H) уже учтён по двойному плану при ⚠️.
        eoh = coverage - demand_H  # coverage = on_hand + inbound_≤H

        # Остаток на момент прибытия первой интранзит-поставки (если она влезает в горизонт)
        if next_eta_mp is not None and (next_eta_mp - t).days <= H:
            days_first = max((next_eta_mp - t).days, 0)
            inbound_first = 0.0
            for it in in_transit:
                if it.sku != x.sku:
                    continue
                eta_mp_i = _eta_to_mp(it, x.lead_time_msk_mp)
                if eta_mp_i < t or eta_mp_i > next_eta_mp:
                    continue
                inbound_first += it.qty
            if stock_status.startswith("⚠️"):
                daily_first = (
                    reduce_plan_to if reduce_plan_to is not None else x.plan_sales_per_day
                )
            else:
                daily_first = x.plan_sales_per_day
            demand_first = daily_first * days_first
            eop_first = (x.stock_ff + x.stock_mp) + float(inbound_first) - float(demand_first)
        else:
            eop_first = None

        reduce_plan_to_display = (
            reduce_plan_to if stock_status.startswith("⚠️") else "–"
        )

        # Комментарий оставляем только как лаконичную метку (для читабельности в Excel)
        comment = "⚙️ dual-plan" if stock_status.startswith("⚠️") else "–"

        recs.append(Recommendation(
            sku=x.sku,
            H_days=H,
            demand_H=demand_H,
            inbound=inbound,
            coverage=coverage,
            target=target,
            shortage=shortage,
            moq_step=x.moq_step,
            order_qty=order_qty,
            stock_status=stock_status,
            reduce_plan_to=reduce_plan_to_display,
            comment=comment,
            algo_version=ALGO_VERSION,
            eoh=eoh,
            eop_first=eop_first,
        ))
    return recs
