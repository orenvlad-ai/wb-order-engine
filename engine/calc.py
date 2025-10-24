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


def _min_stock_with_constant_rate(
    on_hand: float,
    events: List[Tuple[int, int]],
    H: int,
    rate: float,
) -> float:
    stock = on_hand
    min_stock = stock
    prev_day = 0
    for day, qty in events:
        span = max(day - prev_day, 0)
        stock -= rate * span
        min_stock = min(min_stock, stock)
        stock += qty
        prev_day = day
    span = max(H - prev_day, 0)
    stock -= rate * span
    min_stock = min(min_stock, stock)
    return min_stock


def calculate(inputs: List[SkuInput], in_transit: List[InTransitItem]) -> List[Recommendation]:
    t = _today()
    recs: List[Recommendation] = []
    for x in inputs:
        H = _calc_H(x)
        inbound, _ = _inbound_within_H(
            x.sku, in_transit, x.lead_time_msk_mp, H, t
        )
        coverage = x.stock_ff + x.stock_mp + inbound

        events: List[Tuple[int, int]] = []
        for it in in_transit:
            if it.sku != x.sku:
                continue
            eta_mp_i = _eta_to_mp(it, x.lead_time_msk_mp)
            day_offset = (eta_mp_i - t).days
            if 0 <= day_offset <= H:
                events.append((day_offset, it.qty))
        events.sort(key=lambda z: z[0])

        plan = float(x.plan_sales_per_day)
        on_hand = float(x.stock_ff + x.stock_mp)
        oos_threshold = (x.oos_safety_mp_pct / 100.0) * x.safety_stock_mp

        min_stock = _min_stock_with_constant_rate(on_hand, events, H, plan)
        if min_stock < oos_threshold - 1e-9:
            stock_status = "⚠️ Не хватает"
        else:
            stock_status = "Хватает"

        demand_H = float(H) * plan
        target = demand_H + x.safety_stock_mp + x.safety_stock_ff
        shortage = max(0.0, target - coverage)
        order_qty = _order_qty(shortage, x.moq_step)

        # Диагностика остатков по поставкам при текущем плане
        stock_before_1 = stock_after_1 = None
        stock_before_2 = stock_after_2 = None
        stock_before_3 = stock_after_3 = None
        reco_before_1p = reco_before_2p = reco_before_3p = None

        S = on_hand
        prev_day = 0
        def _calc_reco_value(stock_start: float, duration_days: float) -> Optional[int]:
            if duration_days <= 0:
                return None
            if stock_start <= oos_threshold + 1e-9:
                return None
            raw = (stock_start - oos_threshold) / duration_days
            if raw <= 0:
                return 0
            value = math.floor(raw + 1e-9)
            return max(value, 0)

        for idx, (day_offset, qty) in enumerate(events, start=1):
            spend = plan * max(day_offset - prev_day, 0)
            stock_before = S - spend  # «Ост. до XП» может быть отрицательным
            stock_after = max(stock_before, 0.0) + qty  # «Ост. после XП» считаем от нуля
            reco_value = _calc_reco_value(S, max(day_offset - prev_day, 0))
            if idx == 1:
                stock_before_1, stock_after_1 = stock_before, stock_after
                reco_before_1p = reco_value
            elif idx == 2:
                stock_before_2, stock_after_2 = stock_before, stock_after
                reco_before_2p = reco_value
            elif idx == 3:
                stock_before_3, stock_after_3 = stock_before, stock_after
                reco_before_3p = reco_value
            S = stock_after  # следующий участок берём от неотрицательного остатка
            prev_day = day_offset

        tail_duration = max(H - prev_day, 0)
        spend_tail = plan * tail_duration
        stock_before_po = S - spend_tail  # «Ост. до РП» может быть отрицательным
        eoh = stock_before_po  # «Ост. до РП»
        reco_before_po = _calc_reco_value(S, tail_duration) if order_qty > 0 else None
        # «Ост. после РП»: при отсутствии заказа поле пустое
        stock_after_po = (stock_before_po + float(order_qty)) if order_qty > 0 else None
        eop_first = stock_after_1

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
            algo_version=ALGO_VERSION,
            eoh=eoh,
            eop_first=eop_first,
            oos_threshold=oos_threshold,
            reco_before_1p=reco_before_1p,
            stock_before_1=stock_before_1,
            stock_after_1=stock_after_1,
            reco_before_2p=reco_before_2p,
            stock_before_2=stock_before_2,
            stock_after_2=stock_after_2,
            reco_before_3p=reco_before_3p,
            stock_before_3=stock_before_3,
            stock_after_3=stock_after_3,
            reco_before_po=reco_before_po,
            stock_before_po=stock_before_po,
            stock_after_po=stock_after_po,
        ))
    return recs
