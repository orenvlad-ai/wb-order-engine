from datetime import date, timedelta
import math
from typing import Iterable, List, Optional, Tuple

from .models import SkuInput, InTransitItem, Recommendation
from .config import ALGO_VERSION


def _clamp(value: float, lower: float, upper: float) -> float:
    if lower > upper:
        lower, upper = upper, lower
    return max(lower, min(upper, value))


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

        # Собираем все поставки в горизонте H на уровне МП (день от t и qty)
        events: List[Tuple[int, int]] = []
        for it in in_transit:
            if it.sku != x.sku:
                continue
            eta_mp_i = _eta_to_mp(it, x.lead_time_msk_mp)
            d = (eta_mp_i - t).days
            if 0 <= d <= H:
                events.append((d, it.qty))
        events.sort(key=lambda z: z[0])

        plan = x.plan_sales_per_day

        # Проверка на провал к порогу при текущем плане (без снижения)
        def min_stock_with_piecewise(p_before_first: float, p_after_first: float) -> float:
            s = on_hand
            min_s = s
            prev = 0
            if events:
                d1, q1 = events[0]
                L = max(d1 - prev, 0)
                s -= p_before_first * L
                if s < min_s:
                    min_s = s
                s += q1
                prev = d1
                for d, q in events[1:]:
                    L = max(d - prev, 0)
                    s -= p_after_first * L
                    if s < min_s:
                        min_s = s
                    s += q
                    prev = d
                L = max(H - prev, 0)
                s -= p_after_first * L
                if s < min_s:
                    min_s = s
                return min_s
            else:
                s -= p_before_first * H
                return min(s, min_s)

        min_s_no_reduce = min_stock_with_piecewise(plan, plan)

        # Флаг срабатывает, если минимальный прогнозный запас на горизонте ниже порога OOS
        flag_any = min_s_no_reduce < oos_threshold

        # Подбор минимально достаточного снижения плана до первой поставки
        reduce_plan_to: Optional[float] = None
        reduce_plan_to_after: Optional[float] = None
        if flag_any:
            lo, hi = 0.0, float(max(0.0, plan))
            for _ in range(32):
                mid = (lo + hi) / 2.0
                ms = min_stock_with_piecewise(mid, plan)
                if ms >= oos_threshold - 1e-9:
                    lo = mid
                else:
                    hi = mid
            r = math.floor(lo + 1e-9)
            reduce_plan_to = float(min(plan, max(0.0, r)))
            stock_status = "⚠️ Не хватает"
        else:
            stock_status = "Хватает"
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

        if (
            next_eta_mp is not None
            and (next_eta_mp - t).days < H
            and eoh < oos_threshold - 1e-9
            and events
        ):
            plan_before_first = (
                reduce_plan_to if reduce_plan_to is not None else x.plan_sales_per_day
            )
            lo, hi = 0.0, float(max(0.0, plan))
            for _ in range(32):
                mid = (lo + hi) / 2.0
                ms = min_stock_with_piecewise(plan_before_first, mid)
                if ms >= oos_threshold - 1e-9:
                    lo = mid
                else:
                    hi = mid
            r2 = math.floor(lo + 1e-9)
            reduce_plan_to_after = float(min(plan, max(0.0, r2)))

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

        debug_r1_smooth = None
        debug_r2_smooth = None
        debug_d1 = None
        debug_d2 = None
        debug_demand_first = None
        debug_demand_after = None
        debug_eoh_before = None
        debug_eoh_after = None

        if reduce_plan_to is not None or reduce_plan_to_after is not None:
            r1_min = (
                reduce_plan_to
                if reduce_plan_to is not None
                else float(x.plan_sales_per_day)
            )
            r2_min = (
                reduce_plan_to_after
                if reduce_plan_to_after is not None
                else float(x.plan_sales_per_day)
            )

            d1 = float(min(max(days_until_next_inbound, 0.0), H))
            d2 = float(max(H - d1, 0.0))
            horizon = d1 + d2
            if horizon > 1e-9:
                r_avg = (r1_min * d1 + r2_min * d2) / horizon
            else:
                r_avg = r1_min

            r1_smooth = _clamp(r_avg, r1_min, float(x.plan_sales_per_day))
            r2_smooth = _clamp(r_avg, r2_min, float(x.plan_sales_per_day))

            r1_smooth = float(math.floor(r1_smooth + 1e-9))
            r2_smooth = float(math.floor(r2_smooth + 1e-9))

            if r1_smooth < r1_min:
                r1_smooth = float(r1_min)
            if r2_smooth < r2_min:
                r2_smooth = float(r2_min)

            if events:
                while (
                    min_stock_with_piecewise(r1_smooth, r2_smooth)
                    < oos_threshold - 1e-9
                    and r2_smooth > r2_min
                ):
                    new_r2 = max(r2_min, r2_smooth - 1.0)
                    if new_r2 == r2_smooth:
                        break
                    r2_smooth = new_r2

            if reduce_plan_to is not None:
                reduce_plan_to = r1_smooth
            if reduce_plan_to_after is not None:
                reduce_plan_to_after = r2_smooth
            elif r2_smooth < float(x.plan_sales_per_day):
                reduce_plan_to_after = r2_smooth

            # --- ВАЖНО: пересчёты ПОСЛЕ сглаживания ---
            # Длины отрезков
            d1 = float(min(max(days_until_next_inbound, 0.0), H))
            d2 = float(max(H - d1, 0.0))

            # 1) Спрос за горизонт, шт (при наличии поставки: r1 на d1, r2 на d2; иначе r1 на весь H)
            if d1 > 1e-9 and events:
                demand_H = r1_smooth * d1 + r2_smooth * d2
            else:
                demand_H = r1_smooth * float(H)

            # 2) Цель и нехватка
            target = demand_H + x.safety_stock_mp + x.safety_stock_ff
            shortage = max(0.0, target - coverage)

            # 3) Заказ (с округлением до кратности)
            order_qty = _order_qty(shortage, x.moq_step)

            # 4) Остаток после 1-й поставки (если она в горизонте) — пересчитать по r1_smooth
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
                demand_first = r1_smooth * days_first
                eop_first = (x.stock_ff + x.stock_mp) + float(inbound_first) - float(demand_first)
            else:
                eop_first = None

            # 5) Остаток к приходу расчётной партии (конец H) — пересчитать, используя r1_smooth/r2_smooth
            # coverage = on_hand + inbound<=H уже посчитан выше
            eoh = coverage - demand_H
            debug_eoh_before = eoh

            # --- Мгновенная безопасная коррекция r2_smooth по eoh ---
            if eoh < oos_threshold and d2 > 1e-9:
                # Максимально допустимый план после 1-й поставки,
                # при котором запас на конец горизонта не упадёт ниже порога OOS:
                # eoh = coverage - (r1_smooth * d1 + r2 * d2) >= oos_threshold
                # → r2 <= (coverage - oos_threshold - r1_smooth * d1) / d2
                r2_max_safe = (coverage - oos_threshold - r1_smooth * d1) / d2
                r2_smooth = max(r2_min, min(r2_smooth, r2_max_safe))

                # Пересчёт «спрос/цель/нехватка/заказ» под финальные r1/r2
                demand_H = r1_smooth * d1 + r2_smooth * d2
                eoh = coverage - demand_H
                target = demand_H + x.safety_stock_mp + x.safety_stock_ff
                shortage = max(0.0, target - coverage)
                order_qty = _order_qty(shortage, x.moq_step)

                # Зафиксировать обновлённый второй план в выдаче
                if reduce_plan_to_after is not None or r2_smooth < float(x.plan_sales_per_day):
                    reduce_plan_to_after = float(max(r2_min, math.floor(r2_smooth + 1e-9)))

            # --- Если eoh все ещё ниже порога, зажать r1_smooth (до 1-й поставки) ---
            if eoh < oos_threshold and d1 > 1e-9:
                r1_max_safe = (coverage - oos_threshold - r2_smooth * d2) / d1
                r1_smooth = max(r1_min, min(r1_smooth, r1_max_safe))
                r1_smooth = float(math.floor(r1_smooth + 1e-9))

                demand_H = r1_smooth * d1 + r2_smooth * d2
                eoh = coverage - demand_H
                target = demand_H + x.safety_stock_mp + x.safety_stock_ff
                shortage = max(0.0, target - coverage)
                order_qty = _order_qty(shortage, x.moq_step)

                if reduce_plan_to is not None:
                    reduce_plan_to = float(max(r1_min, math.floor(r1_smooth + 1e-9)))

            debug_r1_smooth = r1_smooth
            debug_r2_smooth = r2_smooth
            debug_d1 = d1
            debug_d2 = d2
            debug_demand_first = r1_smooth * d1
            debug_demand_after = r2_smooth * d2
            debug_eoh_after = eoh

        reduce_plan_to_display = (
            reduce_plan_to if stock_status.startswith("⚠️") else "–"
        )
        reduce_plan_to_after_display = (
            reduce_plan_to_after if reduce_plan_to_after is not None else "–"
        )

        # Комментарий: короткая метка dual-plan / "–"
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
            reduce_plan_to_after=reduce_plan_to_after_display,
            comment=comment,
            algo_version=ALGO_VERSION,
            eoh=eoh,
            eop_first=eop_first,
            # ниже — временная диагностика, попадет в Log (см. adapters/excel_io.py)
            debug_r1_smooth=debug_r1_smooth,
            debug_r2_smooth=debug_r2_smooth,
            debug_d1=debug_d1,
            debug_d2=debug_d2,
            debug_demand_first=debug_demand_first,
            debug_demand_after=debug_demand_after,
            debug_eoh_before=debug_eoh_before,
            debug_eoh_after=debug_eoh_after,
        ))
    return recs
