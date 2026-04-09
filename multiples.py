from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import math

import pandas as pd


@dataclass
class MultiplesSnapshot:
    summary_cards: list[dict[str, str]]
    category_tables: dict[str, pd.DataFrame]
    inputs_table: pd.DataFrame
    formula_table: pd.DataFrame
    notes: list[str]


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except Exception:
        return False


def _to_float(value: Any) -> float | None:
    if _is_missing(value):
        return None
    try:
        number = float(value)
    except Exception:
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number


def _safe_div(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator is None or abs(denominator) < 1e-9:
        return None
    return numerator / denominator


def _display_money(value: float | None) -> str:
    if value is None:
        return "n/a"
    return f"{value:,.1f}"


def _display_multiple(value: float | None) -> str:
    if value is None:
        return "n/a"
    return f"{value:.2f}x"


def _display_percent(value: float | None) -> str:
    if value is None:
        return "n/a"
    return f"{value:.1%}"


def _display_price(value: float | None) -> str:
    if value is None:
        return "n/a"
    return f"{value:,.2f}"


def _metric_row(category: str, metric: str, value: float | None, display: str, formula: str, note: str) -> dict[str, str | float | None]:
    return {
        "Category": category,
        "Metric": metric,
        "Value Raw": value,
        "Value": display,
        "Formula": formula,
        "Note": note,
    }


def build_multiples_snapshot(dataset) -> MultiplesSnapshot:
    latest = dataset.latest_values
    overview = dataset.overview
    notes: list[str] = []

    market_cap = _to_float(overview.market_cap_m)
    current_price = _to_float(overview.current_price)
    revenue = _to_float(latest.get("Revenue"))
    net_income = _to_float(latest.get("Net Income"))
    cfo = _to_float(latest.get("CFO (actual, optional)"))
    fcf = _to_float(latest.get("FCF (CFO - Capex)"))
    ebitda = _to_float(latest.get("EBITDA"))
    equity = _to_float(latest.get("Equity"))
    cash = _to_float(latest.get("Cash & Equivalents"))
    dividends = _to_float(latest.get("Dividends"))
    intangibles_and_goodwill = _to_float(latest.get("Intangibles & Goodwill"))
    current_assets = _to_float(latest.get("Current Assets"))
    ppe = _to_float(latest.get("PP&E, net"))
    other_non_current_assets = _to_float(latest.get("Other Non-current Assets"))
    short_term_debt = _to_float(latest.get("Short-term Debt")) or 0.0
    long_term_debt = _to_float(latest.get("Long-term Debt")) or 0.0

    total_debt = short_term_debt + long_term_debt
    net_debt = total_debt - (cash or 0.0)
    enterprise_value = market_cap + total_debt - (cash or 0.0) if market_cap is not None else None
    shares_outstanding = _safe_div(market_cap, current_price) if market_cap is not None and current_price is not None and current_price > 0 else None

    total_assets_parts = [current_assets, ppe, intangibles_and_goodwill, other_non_current_assets]
    total_assets = sum(part for part in total_assets_parts if part is not None) if any(part is not None for part in total_assets_parts) else None

    if intangibles_and_goodwill is None:
        notes.append(
            "Yahoo mapping does not provide a clean standalone goodwill field here, so the app uses the combined `Intangibles & Goodwill` balance when that metric is shown."
        )
    if market_cap is None:
        notes.append("Yahoo did not return a current market cap, so market-based multiples are unavailable.")

    rows = []

    rows.append(
        _metric_row(
            "Market Value",
            "Current price",
            current_price,
            _display_price(current_price),
            "Current Yahoo Finance share price",
            "Live market input from Yahoo Finance.",
        )
    )
    rows.append(
        _metric_row(
            "Market Value",
            "Market cap ($mm)",
            market_cap,
            _display_money(market_cap),
            "Current Yahoo Finance equity market value",
            "Live market input from Yahoo Finance.",
        )
    )
    rows.append(
        _metric_row(
            "Market Value",
            "Enterprise value ($mm)",
            enterprise_value,
            _display_money(enterprise_value),
            "Market cap + total debt - cash",
            "Uses current market cap and latest annual debt and cash.",
        )
    )
    rows.append(
        _metric_row(
            "Market Value",
            "Shares outstanding (derived, mm)",
            shares_outstanding,
            _display_money(shares_outstanding),
            "Market cap / current price",
            "Derived only when both current price and market cap are available.",
        )
    )

    pe = _safe_div(market_cap, net_income) if market_cap is not None and net_income is not None and net_income > 0 else None
    ps = _safe_div(market_cap, revenue) if market_cap is not None and revenue is not None and revenue > 0 else None
    pcf = _safe_div(market_cap, cfo) if market_cap is not None and cfo is not None and cfo > 0 else None
    pb = _safe_div(market_cap, equity) if market_cap is not None and equity is not None and equity > 0 else None
    ev_ebitda = _safe_div(enterprise_value, ebitda) if enterprise_value is not None and ebitda is not None and ebitda > 0 else None
    ev_sales = _safe_div(enterprise_value, revenue) if enterprise_value is not None and revenue is not None and revenue > 0 else None
    earnings_yield = _safe_div(net_income, market_cap) if net_income is not None and market_cap is not None and market_cap > 0 else None
    fcf_yield = _safe_div(fcf, market_cap) if fcf is not None and market_cap is not None and market_cap > 0 else None
    dividend_yield = _safe_div(dividends, market_cap) if dividends is not None and market_cap is not None and market_cap > 0 else None

    rows.extend(
        [
            _metric_row(
                "Market Multiples",
                "P/E",
                pe,
                _display_multiple(pe),
                "Market cap / latest annual net income",
                "Shown only when latest annual net income is positive.",
            ),
            _metric_row(
                "Market Multiples",
                "P/S",
                ps,
                _display_multiple(ps),
                "Market cap / latest annual revenue",
                "Uses current market cap versus the latest annual revenue.",
            ),
            _metric_row(
                "Market Multiples",
                "P/Cash Flow",
                pcf,
                _display_multiple(pcf),
                "Market cap / latest annual CFO",
                "Shown only when latest annual CFO is positive.",
            ),
            _metric_row(
                "Market Multiples",
                "Price / Book",
                pb,
                _display_multiple(pb),
                "Market cap / latest annual equity",
                "Shown only when latest annual equity is positive.",
            ),
            _metric_row(
                "Market Multiples",
                "Enterprise Value / EBITDA",
                ev_ebitda,
                _display_multiple(ev_ebitda),
                "Enterprise value / latest annual EBITDA",
                "Shown only when latest annual EBITDA is positive.",
            ),
            _metric_row(
                "Market Multiples",
                "Enterprise Value / Sales",
                ev_sales,
                _display_multiple(ev_sales),
                "Enterprise value / latest annual revenue",
                "Uses current enterprise value versus the latest annual revenue.",
            ),
            _metric_row(
                "Market Multiples",
                "Earnings Yield %",
                earnings_yield,
                _display_percent(earnings_yield),
                "Latest annual net income / market cap",
                "Inverse of the market-cap-based earnings multiple when earnings are positive.",
            ),
            _metric_row(
                "Market Multiples",
                "FCF Yield %",
                fcf_yield,
                _display_percent(fcf_yield),
                "Latest annual FCF / market cap",
                "Uses latest annual CFO - capex as FCF.",
            ),
            _metric_row(
                "Market Multiples",
                "Dividend Yield %",
                dividend_yield,
                _display_percent(dividend_yield),
                "Latest annual dividends / market cap",
                "Uses latest annual cash dividends from Yahoo annual cash flow data.",
            ),
        ]
    )

    cfo_to_net_income = _safe_div(cfo, net_income)
    dividend_payout = _safe_div(dividends, net_income) if dividends is not None and net_income is not None and net_income > 0 else None
    goodwill_to_assets = _safe_div(intangibles_and_goodwill, total_assets) if intangibles_and_goodwill is not None and total_assets is not None and total_assets > 0 else None
    net_debt_to_market_cap = _safe_div(net_debt, market_cap) if market_cap is not None and market_cap > 0 else None
    cash_to_market_cap = _safe_div(cash, market_cap) if cash is not None and market_cap is not None and market_cap > 0 else None

    rows.extend(
        [
            _metric_row(
                "Quality & Capital Return",
                "CFO / Net Income",
                cfo_to_net_income,
                _display_multiple(cfo_to_net_income),
                "Latest annual CFO / latest annual net income",
                "A cash-conversion check; values above 1.0x indicate CFO exceeds accounting earnings.",
            ),
            _metric_row(
                "Quality & Capital Return",
                "Dividend payout",
                dividend_payout,
                _display_percent(dividend_payout),
                "Latest annual dividends / latest annual net income",
                "Shown only when latest annual net income is positive.",
            ),
            _metric_row(
                "Quality & Capital Return",
                "Net debt / market cap",
                net_debt_to_market_cap,
                _display_percent(net_debt_to_market_cap),
                "Net debt / market cap",
                "Can be negative when the company holds more cash than debt.",
            ),
            _metric_row(
                "Quality & Capital Return",
                "Cash / market cap",
                cash_to_market_cap,
                _display_percent(cash_to_market_cap),
                "Cash and equivalents / market cap",
                "Shows how large the cash balance is versus current equity value.",
            ),
            _metric_row(
                "Quality & Capital Return",
                "Intangibles & goodwill / total assets",
                goodwill_to_assets,
                _display_percent(goodwill_to_assets),
                "Latest annual intangibles & goodwill / total assets",
                "Uses Yahoo's combined intangibles-and-goodwill field because a standalone goodwill field is not mapped here.",
            ),
        ]
    )

    metrics_df = pd.DataFrame(rows)
    category_tables = {
        category: frame[["Metric", "Value", "Formula", "Note"]].reset_index(drop=True)
        for category, frame in metrics_df.groupby("Category", sort=False)
    }

    summary_metric_order = [
        ("P/E", "Market Multiples"),
        ("P/S", "Market Multiples"),
        ("Enterprise Value / EBITDA", "Market Multiples"),
        ("Earnings Yield %", "Market Multiples"),
        ("CFO / Net Income", "Quality & Capital Return"),
    ]
    summary_cards: list[dict[str, str]] = []
    for metric_name, category in summary_metric_order:
        match = metrics_df[(metrics_df["Category"] == category) & (metrics_df["Metric"] == metric_name)]
        if match.empty:
            continue
        summary_cards.append(
            {
                "title": metric_name,
                "value": str(match.iloc[0]["Value"]),
            }
        )

    inputs_rows = [
        {"Input": "Latest fiscal year", "Value": str(dataset.latest_year)},
        {"Input": "Current price", "Value": _display_price(current_price)},
        {"Input": "Market cap ($mm)", "Value": _display_money(market_cap)},
        {"Input": "Revenue ($mm)", "Value": _display_money(revenue)},
        {"Input": "Net income ($mm)", "Value": _display_money(net_income)},
        {"Input": "CFO ($mm)", "Value": _display_money(cfo)},
        {"Input": "FCF ($mm)", "Value": _display_money(fcf)},
        {"Input": "EBITDA ($mm)", "Value": _display_money(ebitda)},
        {"Input": "Equity ($mm)", "Value": _display_money(equity)},
        {"Input": "Cash ($mm)", "Value": _display_money(cash)},
        {"Input": "Short-term debt ($mm)", "Value": _display_money(short_term_debt)},
        {"Input": "Long-term debt ($mm)", "Value": _display_money(long_term_debt)},
        {"Input": "Total debt ($mm)", "Value": _display_money(total_debt)},
        {"Input": "Net debt ($mm)", "Value": _display_money(net_debt)},
        {"Input": "Dividends ($mm)", "Value": _display_money(dividends)},
        {"Input": "Intangibles & goodwill ($mm)", "Value": _display_money(intangibles_and_goodwill)},
        {"Input": "Total assets ($mm)", "Value": _display_money(total_assets)},
    ]
    inputs_table = pd.DataFrame(inputs_rows)

    formula_table = metrics_df[["Category", "Metric", "Formula", "Note"]].copy().reset_index(drop=True)

    notes.insert(
        0,
        f"These multiples use current Yahoo market value inputs versus the latest annual statement year ({dataset.latest_year}). They are not historical market multiples by year."
    )

    return MultiplesSnapshot(
        summary_cards=summary_cards,
        category_tables=category_tables,
        inputs_table=inputs_table,
        formula_table=formula_table,
        notes=notes,
    )
