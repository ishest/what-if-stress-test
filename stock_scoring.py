from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd

from financial_ratios import MILLION, _build_inputs, _safe_div, _to_float


CATEGORY_WEIGHTS = {
    "Growth": 0.30,
    "Profitability": 0.30,
    "Financial Health": 0.20,
    "Valuation": 0.20,
}


@dataclass
class StockScoringModel:
    overall_score: float | None
    recommendation: str
    category_scores: dict[str, float | None]
    category_tables: dict[str, pd.DataFrame]
    summary_cards: list[dict[str, str]]
    inputs_table: pd.DataFrame
    formula_table: pd.DataFrame
    valuation_history_table: pd.DataFrame
    notes: list[str]


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except Exception:
        return False


def _format_number(value: float | None) -> str:
    if value is None or _is_missing(value):
        return "n/a"
    return f"{value:,.2f}"


def _format_pct(value: float | None) -> str:
    if value is None or _is_missing(value):
        return "n/a"
    return f"{value:.1%}"


def _format_score(value: float | None) -> str:
    if value is None or _is_missing(value):
        return "n/a"
    return f"{value:.1f}/100"


def _score_label(score: float | None) -> str:
    if score is None or _is_missing(score):
        return "n/a"
    if score >= 100:
        return "Strong"
    if score >= 75:
        return "Good"
    if score >= 50:
        return "Average"
    return "Weak"


def _score_higher_better(value: float | None, bands: tuple[float, float, float]) -> float | None:
    if value is None or _is_missing(value):
        return None
    if value < bands[0]:
        return 25.0
    if value < bands[1]:
        return 50.0
    if value < bands[2]:
        return 75.0
    return 100.0


def _score_lower_better(value: float | None, bands: tuple[float, float, float]) -> float | None:
    if value is None or _is_missing(value):
        return None
    if value <= bands[0]:
        return 100.0
    if value <= bands[1]:
        return 75.0
    if value <= bands[2]:
        return 50.0
    return 25.0


def _score_range(
    value: float | None,
    strong_range: tuple[float, float],
    good_range: tuple[float, float],
    average_range: tuple[float, float],
) -> float | None:
    if value is None or _is_missing(value):
        return None
    if strong_range[0] <= value <= strong_range[1]:
        return 100.0
    if good_range[0] <= value <= good_range[1]:
        return 75.0
    if average_range[0] <= value <= average_range[1]:
        return 50.0
    return 25.0


def _score_relative_multiple(value: float | None, bands: tuple[float, float, float]) -> float | None:
    return _score_lower_better(value, bands)


def _cagr(start_value: float | None, end_value: float | None, periods: int) -> float | None:
    if periods <= 0 or start_value is None or end_value is None:
        return None
    if start_value <= 0 or end_value <= 0:
        return None
    return (end_value / start_value) ** (1 / periods) - 1


def _sort_year_index(index: pd.Index) -> list[str]:
    def _key(label: Any) -> tuple[int, str]:
        text = str(label)
        try:
            return int(text), text
        except Exception:
            return 0, text

    return [str(label) for label in sorted(index, key=_key)]


def _recommendation(score: float | None) -> str:
    if score is None or _is_missing(score):
        return "N/A"
    if score >= 85:
        return "Strong Buy"
    if score >= 70:
        return "Buy"
    if score >= 55:
        return "Hold"
    if score >= 40:
        return "Sell"
    return "Strong Sell"


def _weighted_average(scores: dict[str, float | None]) -> float | None:
    weighted_total = 0.0
    used_weight = 0.0
    for category, weight in CATEGORY_WEIGHTS.items():
        score = scores.get(category)
        if score is None or _is_missing(score):
            continue
        weighted_total += score * weight
        used_weight += weight
    if used_weight <= 0:
        return None
    return weighted_total / used_weight


def _metric_row(
    category: str,
    metric: str,
    value: float | None,
    display_value: str,
    score: float | None,
    formula: str,
    threshold_logic: str,
    note: str,
) -> dict[str, str | float | None]:
    return {
        "Category": category,
        "Metric": metric,
        "Value Raw": value,
        "Value": display_value,
        "Score Raw": score,
        "Score": _format_score(score),
        "Band": _score_label(score),
        "Formula": formula,
        "Threshold logic": threshold_logic,
        "Note": note,
    }


def build_stock_scoring_model(dataset) -> StockScoringModel:
    notes: list[str] = []
    inputs = _build_inputs(dataset)
    ordered_years = _sort_year_index(inputs.index)
    inputs = inputs.loc[ordered_years]
    latest_year = ordered_years[-1] if ordered_years else None
    trailing_years = ordered_years[-3:] if len(ordered_years) >= 3 else ordered_years
    growth_start_year = trailing_years[0] if trailing_years else None
    periods = max(len(trailing_years) - 1, 0)

    latest = inputs.loc[latest_year] if latest_year is not None else pd.Series(dtype=float)
    start = inputs.loc[growth_start_year] if growth_start_year is not None else pd.Series(dtype=float)

    current_price = _to_float(dataset.overview.current_price)
    market_cap_m = _to_float(dataset.overview.market_cap_m)

    revenue_cagr = _cagr(_to_float(start.get("Revenue")), _to_float(latest.get("Revenue")), periods) if periods else None
    earnings_cagr = _cagr(_to_float(start.get("Net Income")), _to_float(latest.get("Net Income")), periods) if periods else None

    operating_margin_latest = _safe_div(_to_float(latest.get("EBIT")), _to_float(latest.get("Revenue")))
    net_margin_latest = _safe_div(_to_float(latest.get("Net Income")), _to_float(latest.get("Revenue")))
    operating_margin_start = _safe_div(_to_float(start.get("EBIT")), _to_float(start.get("Revenue")))
    margin_trend = (
        operating_margin_latest - operating_margin_start
        if operating_margin_latest is not None and operating_margin_start is not None
        else None
    )
    roic = _safe_div(_to_float(latest.get("NOPAT")), _to_float(latest.get("Average Invested Capital")))

    debt_to_equity = _safe_div(_to_float(latest.get("Total Debt")), _to_float(latest.get("Equity")))
    current_ratio = _safe_div(_to_float(latest.get("Current Assets")), _to_float(latest.get("Current Liabilities")))
    quick_assets = None
    if _to_float(latest.get("Current Assets")) is not None and _to_float(latest.get("Inventory")) is not None:
        quick_assets = _to_float(latest.get("Current Assets")) - _to_float(latest.get("Inventory"))
    quick_ratio = _safe_div(quick_assets, _to_float(latest.get("Current Liabilities")))
    interest_coverage = _safe_div(_to_float(latest.get("EBIT")), _to_float(latest.get("Interest Expense")))

    latest_net_income = _to_float(latest.get("Net Income"))
    latest_revenue = _to_float(latest.get("Revenue"))
    latest_ebitda = _to_float(latest.get("EBITDA"))
    latest_cash = _to_float(latest.get("Cash"))
    latest_total_debt = _to_float(latest.get("Total Debt"))
    enterprise_value_m = (
        market_cap_m + latest_total_debt - latest_cash
        if market_cap_m is not None and latest_total_debt is not None and latest_cash is not None
        else None
    )
    pe_current = _safe_div(market_cap_m, latest_net_income) if latest_net_income and latest_net_income > 0 else None
    ps_current = _safe_div(market_cap_m, latest_revenue) if latest_revenue and latest_revenue > 0 else None
    ev_ebitda = _safe_div(enterprise_value_m, latest_ebitda) if latest_ebitda and latest_ebitda > 0 else None
    earnings_yield = _safe_div(latest_net_income, market_cap_m) if latest_net_income and market_cap_m and market_cap_m > 0 else None

    if market_cap_m is None or current_price is None:
        notes.append("Yahoo did not return current market cap or price cleanly, so some valuation metrics may be unavailable.")

    raw = dataset.annual_raw.copy()
    raw["asOfDate"] = pd.to_datetime(raw["asOfDate"], errors="coerce")
    raw["Year"] = raw["asOfDate"].dt.year.astype("Int64").astype(str)
    raw = raw.sort_values("asOfDate")

    valuation_history_rows: list[dict[str, str | float | None]] = []
    for year in trailing_years:
        matched = raw[raw["Year"] == str(year)]
        if matched.empty:
            continue
        raw_row = matched.iloc[-1]
        fiscal_date = pd.to_datetime(raw_row["asOfDate"], errors="coerce")
        year_market_cap_raw = _to_float(raw_row.get("MarketCap"))
        year_revenue = _to_float(raw_row.get("TotalRevenue"))
        year_net_income = _to_float(raw_row.get("NetIncome"))
        year_market_cap = year_market_cap_raw / MILLION if year_market_cap_raw is not None else None
        year_pe = _safe_div(year_market_cap, year_net_income / MILLION) if year_market_cap is not None and year_net_income and year_net_income > 0 else None
        year_ps = _safe_div(year_market_cap, year_revenue / MILLION) if year_market_cap is not None and year_revenue and year_revenue > 0 else None
        valuation_history_rows.append(
            {
                "Year": year,
                "Fiscal date": fiscal_date.strftime("%Y-%m-%d") if not pd.isna(fiscal_date) else "n/a",
                "Historic market cap ($mm)": year_market_cap,
                "Historic P/E": year_pe,
                "Historic P/S": year_ps,
            }
        )

    valuation_history_table = pd.DataFrame(valuation_history_rows)
    if not valuation_history_table.empty:
        pe_history_series = pd.to_numeric(valuation_history_table["Historic P/E"], errors="coerce").dropna()
        ps_history_series = pd.to_numeric(valuation_history_table["Historic P/S"], errors="coerce").dropna()
        pe_history_avg = _to_float(pe_history_series.mean()) if not pe_history_series.empty else None
        ps_history_avg = _to_float(ps_history_series.mean()) if not ps_history_series.empty else None
    else:
        pe_history_avg = None
        ps_history_avg = None

    pe_vs_history = _safe_div(pe_current, pe_history_avg) if pe_current is not None and pe_history_avg is not None and pe_history_avg > 0 else None
    ps_vs_history = _safe_div(ps_current, ps_history_avg) if ps_current is not None and ps_history_avg is not None and ps_history_avg > 0 else None

    if valuation_history_table.empty:
        notes.append("Historical valuation context could not be built from Yahoo annual market-cap fields for the last three fiscal years.")

    metric_rows = [
        _metric_row(
            "Growth",
            "Revenue CAGR",
            revenue_cagr,
            _format_pct(revenue_cagr),
            _score_higher_better(revenue_cagr, (0.00, 0.05, 0.10)),
            "((Revenue Y3 / Revenue Y1)^(1/periods)) - 1",
            "<0%=25 | 0-5%=50 | 5-10%=75 | >=10%=100",
            "Uses the latest three Yahoo annual periods when available.",
        ),
        _metric_row(
            "Growth",
            "Earnings CAGR",
            earnings_cagr,
            _format_pct(earnings_cagr),
            _score_higher_better(earnings_cagr, (0.00, 0.06, 0.12)),
            "((Net income Y3 / Net income Y1)^(1/periods)) - 1",
            "<0%=25 | 0-6%=50 | 6-12%=75 | >=12%=100",
            "Returns `n/a` if the starting or ending year is non-positive.",
        ),
        _metric_row(
            "Profitability",
            "Operating Margin",
            operating_margin_latest,
            _format_pct(operating_margin_latest),
            _score_higher_better(operating_margin_latest, (0.05, 0.10, 0.20)),
            "EBIT / revenue (latest year)",
            "<5%=25 | 5-10%=50 | 10-20%=75 | >=20%=100",
            "Latest annual operating profitability from Yahoo statements.",
        ),
        _metric_row(
            "Profitability",
            "Net Margin",
            net_margin_latest,
            _format_pct(net_margin_latest),
            _score_higher_better(net_margin_latest, (0.03, 0.08, 0.15)),
            "Net income / revenue (latest year)",
            "<3%=25 | 3-8%=50 | 8-15%=75 | >=15%=100",
            "Latest annual net profitability from Yahoo statements.",
        ),
        _metric_row(
            "Profitability",
            "Margin Trend",
            margin_trend,
            _format_pct(margin_trend),
            _score_higher_better(margin_trend, (0.00, 0.01, 0.03)),
            "Operating margin latest year - operating margin first comparison year",
            "<0bps=25 | 0-100bps=50 | 100-300bps=75 | >=300bps=100",
            "Captures whether operating profitability is improving or deteriorating over the three-year window.",
        ),
        _metric_row(
            "Profitability",
            "ROIC",
            roic,
            _format_pct(roic),
            _score_higher_better(roic, (0.05, 0.10, 0.15)),
            "NOPAT / average invested capital",
            "<5%=25 | 5-10%=50 | 10-15%=75 | >=15%=100",
            "NOPAT uses Yahoo taxes and pretax income where available.",
        ),
        _metric_row(
            "Financial Health",
            "Debt-to-Equity",
            debt_to_equity,
            _format_number(debt_to_equity),
            _score_lower_better(debt_to_equity, (0.50, 1.00, 2.00)),
            "Total debt / equity",
            "<=0.5x=100 | <=1.0x=75 | <=2.0x=50 | >2.0x=25",
            "Lower leverage scores better.",
        ),
        _metric_row(
            "Financial Health",
            "Current Ratio",
            current_ratio,
            _format_number(current_ratio),
            _score_higher_better(current_ratio, (1.00, 1.30, 1.80)),
            "Current assets / current liabilities",
            "<1.0x=25 | 1.0-1.3x=50 | 1.3-1.8x=75 | >=1.8x=100",
            "Uses the latest annual current asset and liability balances.",
        ),
        _metric_row(
            "Financial Health",
            "Quick Ratio",
            quick_ratio,
            _format_number(quick_ratio),
            _score_higher_better(quick_ratio, (0.70, 1.00, 1.50)),
            "(Current assets - inventory) / current liabilities",
            "<0.7x=25 | 0.7-1.0x=50 | 1.0-1.5x=75 | >=1.5x=100",
            "Measures near-cash liquidity without relying on inventory monetization.",
        ),
        _metric_row(
            "Financial Health",
            "Interest Coverage",
            interest_coverage,
            _format_number(interest_coverage),
            _score_higher_better(interest_coverage, (3.00, 6.00, 12.00)),
            "EBIT / interest expense",
            "<3x=25 | 3-6x=50 | 6-12x=75 | >=12x=100",
            "Uses annual EBIT and interest expense from Yahoo statements.",
        ),
        _metric_row(
            "Valuation",
            "P/E",
            pe_current,
            _format_number(pe_current),
            _score_lower_better(pe_current, (15.0, 22.0, 30.0)),
            "Current market cap / latest annual net income",
            "<=15x=100 | <=22x=75 | <=30x=50 | >30x=25",
            "Shown only when latest annual net income is positive.",
        ),
        _metric_row(
            "Valuation",
            "P/S",
            ps_current,
            _format_number(ps_current),
            _score_lower_better(ps_current, (2.0, 4.0, 7.0)),
            "Current market cap / latest annual revenue",
            "<=2x=100 | <=4x=75 | <=7x=50 | >7x=25",
            "Uses current market cap against latest annual sales.",
        ),
        _metric_row(
            "Valuation",
            "Enterprise Value / EBITDA",
            ev_ebitda,
            _format_number(ev_ebitda),
            _score_lower_better(ev_ebitda, (8.0, 12.0, 18.0)),
            "Enterprise value / latest annual EBITDA",
            "<=8x=100 | <=12x=75 | <=18x=50 | >18x=25",
            "Shown only when EBITDA is positive.",
        ),
        _metric_row(
            "Valuation",
            "Earnings Yield",
            earnings_yield,
            _format_pct(earnings_yield),
            _score_higher_better(earnings_yield, (0.03, 0.05, 0.08)),
            "Latest annual net income / current market cap",
            "<3%=25 | 3-5%=50 | 5-8%=75 | >=8%=100",
            "Inverse of the market-cap-based earnings multiple.",
        ),
        _metric_row(
            "Valuation",
            "P/E vs 3Y Avg",
            pe_vs_history,
            _format_number(pe_vs_history),
            _score_relative_multiple(pe_vs_history, (0.85, 1.00, 1.20)),
            "Current P/E / average historic P/E from the last three fiscal years",
            "<=0.85x=100 | <=1.0x=75 | <=1.2x=50 | >1.2x=25",
            "Lower means the stock is cheaper than its own recent Yahoo annual valuation history.",
        ),
        _metric_row(
            "Valuation",
            "P/S vs 3Y Avg",
            ps_vs_history,
            _format_number(ps_vs_history),
            _score_relative_multiple(ps_vs_history, (0.85, 1.00, 1.20)),
            "Current P/S / average historic P/S from the last three fiscal years",
            "<=0.85x=100 | <=1.0x=75 | <=1.2x=50 | >1.2x=25",
            "Lower means the stock is cheaper than its own recent Yahoo annual sales valuation history.",
        ),
    ]

    metrics_df = pd.DataFrame(metric_rows)
    category_tables: dict[str, pd.DataFrame] = {}
    category_scores: dict[str, float | None] = {}
    for category in ["Growth", "Profitability", "Financial Health", "Valuation"]:
        category_df = metrics_df[metrics_df["Category"] == category][
            ["Metric", "Value", "Score", "Band", "Threshold logic", "Note"]
        ].reset_index(drop=True)
        category_tables[category] = category_df
        raw_scores = pd.to_numeric(metrics_df.loc[metrics_df["Category"] == category, "Score Raw"], errors="coerce").dropna()
        category_scores[category] = float(raw_scores.mean()) if not raw_scores.empty else None

    overall_score = _weighted_average(category_scores)
    recommendation = _recommendation(overall_score)

    summary_cards = [
        {"title": "Overall Score", "value": _format_score(overall_score)},
        {"title": "Recommendation", "value": recommendation},
        {"title": "Growth", "value": _format_score(category_scores.get("Growth"))},
        {"title": "Profitability", "value": _format_score(category_scores.get("Profitability"))},
        {"title": "Financial Health", "value": _format_score(category_scores.get("Financial Health"))},
        {"title": "Valuation", "value": _format_score(category_scores.get("Valuation"))},
    ]

    inputs_rows = [
        {"Input": "Ticker", "Value": dataset.overview.ticker},
        {"Input": "Company", "Value": dataset.overview.long_name or dataset.overview.short_name},
        {"Input": "Sector", "Value": dataset.overview.sector},
        {"Input": "Industry", "Value": dataset.overview.industry},
        {"Input": "Latest annual year", "Value": latest_year or "n/a"},
        {"Input": "Growth window", "Value": " -> ".join(trailing_years) if trailing_years else "n/a"},
        {"Input": "Revenue start ($mm)", "Value": _format_number(_to_float(start.get("Revenue")))},
        {"Input": "Revenue latest ($mm)", "Value": _format_number(_to_float(latest.get("Revenue")))},
        {"Input": "Net income start ($mm)", "Value": _format_number(_to_float(start.get("Net Income")))},
        {"Input": "Net income latest ($mm)", "Value": _format_number(_to_float(latest.get("Net Income")))},
        {"Input": "EBIT latest ($mm)", "Value": _format_number(_to_float(latest.get("EBIT")))},
        {"Input": "EBITDA latest ($mm)", "Value": _format_number(_to_float(latest.get("EBITDA")))},
        {"Input": "Current price", "Value": _format_number(current_price)},
        {"Input": "Market cap ($mm)", "Value": _format_number(market_cap_m)},
        {"Input": "Enterprise value ($mm)", "Value": _format_number(enterprise_value_m)},
        {"Input": "Historic average P/E (3Y)", "Value": _format_number(pe_history_avg)},
        {"Input": "Historic average P/S (3Y)", "Value": _format_number(ps_history_avg)},
    ]
    inputs_table = pd.DataFrame(inputs_rows)

    formula_rows = [
        {
            "Category": category,
            "Metric": metric,
            "Formula": formula,
            "Threshold logic": threshold_logic,
            "Category weight": f"{CATEGORY_WEIGHTS[category]:.0%}",
            "Note": note,
        }
        for category, metric, formula, threshold_logic, note in metrics_df[
            ["Category", "Metric", "Formula", "Threshold logic", "Note"]
        ].itertuples(index=False, name=None)
    ]
    formula_rows.extend(
        [
            {
                "Category": "Overall",
                "Metric": "Weighted score",
                "Formula": "Weighted average of available category scores",
                "Threshold logic": "Growth 30% | Profitability 30% | Financial Health 20% | Valuation 20%",
                "Category weight": "100%",
                "Note": "If a category is `n/a`, its weight is excluded and the remaining weights are normalized.",
            },
            {
                "Category": "Overall",
                "Metric": "Recommendation",
                "Formula": "Map the overall score to the recommendation band",
                "Threshold logic": "85-100 Strong Buy | 70-84 Buy | 55-69 Hold | 40-54 Sell | <40 Strong Sell",
                "Category weight": "n/a",
                "Note": "This is a rule-based model output, not investment advice.",
            },
        ]
    )
    formula_table = pd.DataFrame(formula_rows)

    if current_price is None:
        notes.append("Current Yahoo share price is missing, so market-based valuation scoring is partially unavailable.")
    if pe_history_avg is None or ps_history_avg is None:
        notes.append("At least one historical valuation comparison metric is unavailable because annual market-cap, revenue, or net income data was incomplete.")
    notes.append("This implementation uses only Yahoo Finance data for the individual ticker. It does not fabricate industry benchmark data.")

    if not valuation_history_table.empty:
        valuation_history_table = valuation_history_table.copy()
        valuation_history_table["Historic market cap ($mm)"] = valuation_history_table["Historic market cap ($mm)"].map(_format_number)
        valuation_history_table["Historic P/E"] = valuation_history_table["Historic P/E"].map(_format_number)
        valuation_history_table["Historic P/S"] = valuation_history_table["Historic P/S"].map(_format_number)

    return StockScoringModel(
        overall_score=overall_score,
        recommendation=recommendation,
        category_scores=category_scores,
        category_tables=category_tables,
        summary_cards=summary_cards,
        inputs_table=inputs_table,
        formula_table=formula_table,
        valuation_history_table=valuation_history_table,
        notes=notes,
    )
