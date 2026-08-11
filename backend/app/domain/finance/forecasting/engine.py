from dataclasses import dataclass
from decimal import Decimal
from typing import Sequence


@dataclass(frozen=True)
class ForecastPoint:
    period: str
    base: Decimal
    downside: Decimal
    upside: Decimal


def build_run_rate_forecast(history: Sequence[tuple[str, Decimal]], future_periods: Sequence[str], *, downside_factor=Decimal("0.90"), upside_factor=Decimal("1.10"), recent_months=3) -> tuple[ForecastPoint,...]:
    values=[v for _,v in history]
    if not values: return tuple()
    sample=values[-recent_months:] if len(values)>=recent_months else values
    base=sum(sample,Decimal("0"))/Decimal(len(sample))
    return tuple(ForecastPoint(p,base,base*downside_factor,base*upside_factor) for p in future_periods)


def build_trend_forecast(history: Sequence[tuple[str, Decimal]], future_periods: Sequence[str], *, downside_factor=Decimal("0.90"), upside_factor=Decimal("1.10")) -> tuple[ForecastPoint,...]:
    values=[v for _,v in history]
    if not values:return tuple()
    slope=(values[-1]-values[0])/Decimal(max(1,len(values)-1))
    result=[]
    for i,p in enumerate(future_periods,1):
        base=values[-1]+slope*Decimal(i); result.append(ForecastPoint(p,base,base*downside_factor,base*upside_factor))
    return tuple(result)
