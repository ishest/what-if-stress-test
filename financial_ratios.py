from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import math

import pandas as pd


MILLION = 1_000_000.0
CATEGORY_ORDER = ["Total score", "Profitability", "Liquidity", "Credit risk", "Activity", "Market multiples"]


@dataclass
class RatioScorecard:
    summary_scores: dict[str, float | None]
    category_tables: dict[str, pd.DataFrame]
    derived_table: pd.DataFrame
    formula_table: pd.DataFrame
    notes: list[str]


RATIO_DEFINITIONS = [
    {
        "category": "Profitability",
        "label": "ROIC",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "NOPAT / average invested capital",
    },
    {
        "category": "Profitability",
        "label": "Gross margin",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "Gross profit / revenue",
    },
    {
        "category": "Profitability",
        "label": "Operating margin",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "EBIT / revenue",
    },
    {
        "category": "Profitability",
        "label": "Net margin",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "Net income / revenue",
    },
    {
        "category": "Profitability",
        "label": "ROE",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "Net income / average equity",
    },
    {
        "category": "Profitability",
        "label": "ROA",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "Net income / average assets",
    },
    {
        "category": "Liquidity",
        "label": "Current ratio",
        "mark": "1-2",
        "kind": "range",
        "target": (1.0, 2.0),
        "format": "number",
        "formula": "Current assets / current liabilities",
    },
    {
        "category": "Liquidity",
        "label": "Quick ratio",
        "mark": "1",
        "kind": "threshold_min",
        "target": 1.0,
        "format": "number",
        "formula": "(Current assets - inventory) / current liabilities",
    },
    {
        "category": "Liquidity",
        "label": "Cash ratio",
        "mark": "0.20",
        "kind": "threshold_min",
        "target": 0.20,
        "format": "number",
        "formula": "Cash and equivalents / current liabilities",
    },
    {
        "category": "Liquidity",
        "label": "Receivable turnover",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "Revenue / average receivables",
    },
    {
        "category": "Liquidity",
        "label": "Cash conversion cycle",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "Days sales outstanding + days inventory - days payables",
    },
    {
        "category": "Credit risk",
        "label": "Altman Z score",
        "mark": ">2.99",
        "kind": "threshold_min",
        "target": 2.99,
        "format": "number",
        "formula": "1.2*(WC/TA) + 1.4*(RE/TA) + 3.3*(EBIT/TA) + 0.6*(MVE/TL) + 1.0*(Sales/TA)",
    },
    {
        "category": "Credit risk",
        "label": "Net Debt/EBITDA",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "(Short-term debt + long-term debt - cash) / EBITDA",
    },
    {
        "category": "Credit risk",
        "label": "Debt to Equity",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "Total debt / equity",
    },
    {
        "category": "Credit risk",
        "label": "Debt to Capital",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "Total debt / (total debt + equity)",
    },
    {
        "category": "Credit risk",
        "label": "Debt to Assets",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "Total debt / total assets",
    },
    {
        "category": "Credit risk",
        "label": "Financial leverage",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "Total liabilities / total assets",
    },
    {
        "category": "Credit risk",
        "label": "EBITDA to Interest",
        "mark": "max",
        "kind": "max",
        "format": "number",
        "formula": "EBITDA / interest expense",
    },
    {
        "category": "Activity",
        "label": "Days of sales",
        "mark": "avg",
        "kind": "avg",
        "format": "number",
        "formula": "Average receivables / revenue * 365",
    },
    {
        "category": "Activity",
        "label": "Days of inventory",
        "mark": "avg",
        "kind": "avg",
        "format": "number",
        "formula": "Average inventory / COGS * 365",
    },
    {
        "category": "Activity",
        "label": "Inventory turnover",
        "mark": "avg",
        "kind": "avg",
        "format": "number",
        "formula": "COGS / average inventory",
    },
    {
        "category": "Activity",
        "label": "Payable turnover",
        "mark": "avg",
        "kind": "avg",
        "format": "number",
        "formula": "COGS / average accounts payable",
    },
    {
        "category": "Activity",
        "label": "Days of payables",
        "mark": "avg",
        "kind": "avg",
        "format": "number",
        "formula": "Average accounts payable / COGS * 365",
    },
    {
        "category": "Activity",
        "label": "Asset turnover",
        "mark": "avg",
        "kind": "avg",
        "format": "number",
        "formula": "Revenue / average assets",
    },
    {
        "category": "Activity",
        "label": "Leverage ratio",
        "mark": "min",
        "kind": "min",
        "format": "number",
        "formula": "Total assets / equity",
    },
    {
        "category": "Market multiples",
        "label": "P/E",
        "mark": "<= 25x",
        "kind": "threshold_max",
        "target": 25.0,
        "format": "number",
        "formula": "Market cap / net income",
    },
    {
        "category": "Market multiples",
        "label": "P/S",
        "mark": "<= 6x",
        "kind": "threshold_max",
        "target": 6.0,
        "format": "number",
        "formula": "Market cap / revenue",
    },
    {
        "category": "Market multiples",
        "label": "P/Cash Flow",
        "mark": "<= 20x",
        "kind": "threshold_max",
        "target": 20.0,
        "format": "number",
        "formula": "Market cap / CFO",
    },
    {
        "category": "Market multiples",
        "label": "Price / Book",
        "mark": "<= 5x",
        "kind": "threshold_max",
        "target": 5.0,
        "format": "number",
        "formula": "Market cap / equity",
    },
    {
        "category": "Market multiples",
        "label": "Enterprise Value / EBITDA",
        "mark": "<= 15x",
        "kind": "threshold_max",
        "target": 15.0,
        "format": "number",
        "formula": "Enterprise value / EBITDA",
    },
    {
        "category": "Market multiples",
        "label": "Enterprise Value / Sales",
        "mark": "<= 6x",
        "kind": "threshold_max",
        "target": 6.0,
        "format": "number",
        "formula": "Enterprise value / revenue",
    },
]


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


def _to_millions(value: Any) -> float | None:
    number = _to_float(value)
    if number is None:
        return None
    return number / MILLION


def _first_present(row: pd.Series | None, columns: list[str]) -> float | None:
    if row is None:
        return None
    for column in columns:
        value = _to_millions(row.get(column))
        if value is not None:
            return value
    return None


def _safe_div(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator is None or abs(denominator) < 1e-9:
        return None
    return numerator / denominator


def _average_two(current: float | None, previous: float | None) -> float | None:
    if current is None and previous is None:
        return None
    if current is None:
        return previous
    if previous is None:
        return current
    return (current + previous) / 2


def _format_value(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:,.2f}"


def stars_text(score: float | None) -> str:
    if score is None or pd.isna(score):
        return "n/a"
    rounded = int(round(max(0.0, min(5.0, score))))
    return "★" * rounded + "☆" * (5 - rounded)


def _score_from_rank(rank: float) -> int:
    if rank >= 0.9:
        return 5
    if rank >= 0.7:
        return 4
    if rank >= 0.5:
        return 3
    if rank >= 0.3:
        return 2
    if rank > 0:
        return 1
    return 0


def _score_threshold_min(value: float | None, target: float) -> int | None:
    if value is None or pd.isna(value):
        return None
    if value <= 0:
        return 0
    ratio = value / target if target > 0 else 1.0
    if ratio >= 1.5:
        return 5
    if ratio >= 1.25:
        return 4
    if ratio >= 1.0:
        return 3
    if ratio >= 0.75:
        return 2
    return 1


def _score_range(value: float | None, low: float, high: float) -> int | None:
    if value is None or pd.isna(value):
        return None
    if value <= 0:
        return 0
    if low <= value <= high:
        return 5
    if value < low:
        ratio = value / low
        if ratio >= 0.8:
            return 4
        if ratio >= 0.6:
            return 3
        if ratio >= 0.4:
            return 2
        return 1

    if value <= high * 1.2:
        return 4
    if value <= high * 1.5:
        return 3
    if value <= high * 2.0:
        return 2
    return 1


def _score_threshold_max(value: float | None, target: float) -> int | None:
    if value is None or pd.isna(value):
        return None
    if value <= 0:
        return 0
    ratio = value / target if target > 0 else 1.0
    if ratio <= 0.50:
        return 5
    if ratio <= 0.75:
        return 4
    if ratio <= 1.00:
        return 3
    if ratio <= 1.50:
        return 2
    if ratio <= 2.00:
        return 1
    return 0


def _score_series(series: pd.Series, kind: str, target: Any = None) -> int | None:
    clean = pd.to_numeric(series, errors="coerce").dropna()
    if clean.empty:
        return None
    latest = float(clean.iloc[-1])

    if kind == "threshold_min":
        return _score_threshold_min(latest, float(target))
    if kind == "threshold_max":
        return _score_threshold_max(latest, float(target))
    if kind == "range":
        low, high = target
        return _score_range(latest, float(low), float(high))
    if len(clean) == 1:
        return 3 if latest > 0 else 0

    if kind == "max":
        rank = float((clean <= latest).mean())
        if latest <= 0 and clean.max() <= 0:
            return 0
        return _score_from_rank(rank)
    if kind == "min":
        rank = float((clean >= latest).mean())
        if latest <= 0:
            return 5
        return _score_from_rank(rank)
    if kind == "avg":
        median = float(clean.median())
        distances = (clean - median).abs()
        latest_distance = float(abs(latest - median))
        rank = float((distances >= latest_distance).mean())
        return _score_from_rank(rank)

    return None


def _build_inputs(dataset) -> pd.DataFrame:
    financials = dataset.financials
    annual_raw = dataset.annual_raw.copy()
    annual_raw["Fiscal Year"] = annual_raw["Fiscal Year"].astype(str)
    annual_raw = annual_raw.drop_duplicates(subset=["Fiscal Year"], keep="last").set_index("Fiscal Year")

    years = list(financials.columns)
    records: dict[str, dict[str, float | None]] = {}
    latest_year = str(years[-1])

    for year in years:
        raw_row = annual_raw.loc[year] if year in annual_raw.index else None
        revenue = _to_float(financials.at["Revenue", year])
        cogs = _to_float(financials.at["COGS", year])
        gross_profit = _to_float(financials.at["Gross Profit", year])
        ebitda = _to_float(financials.at["EBITDA", year])
        ebit = _to_float(financials.at["EBIT", year])
        taxes = _to_float(financials.at["Taxes", year])
        net_income = _to_float(financials.at["Net Income", year])
        cash = _to_float(financials.at["Cash & Equivalents", year])
        receivables = _to_float(financials.at["Accounts Receivable", year])
        inventory = _to_float(financials.at["Inventory", year])
        other_current_assets = _to_float(financials.at["Other Current Assets", year]) or 0.0
        ppe = _to_float(financials.at["PP&E, net", year])
        intangibles = _to_float(financials.at["Intangibles & Goodwill", year]) or 0.0
        other_non_current_assets = _to_float(financials.at["Other Non-current Assets", year]) or 0.0
        short_term_debt = _to_float(financials.at["Short-term Debt", year]) or 0.0
        payables = _to_float(financials.at["Accounts Payable", year])
        other_current_liabilities = _to_float(financials.at["Other Current Liabilities", year]) or 0.0
        long_term_debt = _to_float(financials.at["Long-term Debt", year]) or 0.0
        other_non_current_liabilities = _to_float(financials.at["Other Non-current Liabilities", year]) or 0.0
        equity = _to_float(financials.at["Equity", year])
        interest_expense = _to_float(financials.at["Interest Expense", year])
        cfo = _to_float(financials.at["CFO (actual, optional)", year])

        current_assets = None
        if None not in (cash, receivables, inventory):
            current_assets = cash + receivables + inventory + other_current_assets

        current_liabilities = None
        if payables is not None:
            current_liabilities = short_term_debt + payables + other_current_liabilities

        derived_total_assets = None
        if None not in (cash, receivables, inventory, ppe):
            derived_total_assets = cash + receivables + inventory + other_current_assets + ppe + intangibles + other_non_current_assets

        derived_total_liabilities = None
        if payables is not None:
            derived_total_liabilities = short_term_debt + payables + other_current_liabilities + long_term_debt + other_non_current_liabilities

        total_assets = _first_present(raw_row, ["TotalAssets"]) or derived_total_assets
        total_liabilities = _first_present(raw_row, ["TotalLiabilitiesNetMinorityInterest", "TotalLiabilities"]) or derived_total_liabilities
        retained_earnings = _first_present(raw_row, ["RetainedEarnings"])
        market_cap = _first_present(raw_row, ["MarketCap"])
        if market_cap is None and year == latest_year:
            market_cap = dataset.overview.market_cap_m

        total_debt = short_term_debt + long_term_debt
        net_debt = total_debt - (cash or 0.0) if cash is not None else None
        enterprise_value = market_cap + total_debt - (cash or 0.0) if market_cap is not None else None
        invested_capital = total_debt + equity - cash if None not in (equity, cash) else None
        working_capital = current_assets - current_liabilities if None not in (current_assets, current_liabilities) else None
        records[year] = {
            "Revenue": revenue,
            "COGS": cogs,
            "Gross Profit": gross_profit,
            "EBITDA": ebitda,
            "EBIT": ebit,
            "Pretax Income": _to_float(financials.at["Pretax Income", year]),
            "Taxes": taxes,
            "Net Income": net_income,
            "CFO": cfo,
            "Cash": cash,
            "Receivables": receivables,
            "Inventory": inventory,
            "Other Current Assets": other_current_assets,
            "PP&E": ppe,
            "Intangibles": intangibles,
            "Other Non-current Assets": other_non_current_assets,
            "Short-term Debt": short_term_debt,
            "Payables": payables,
            "Other Current Liabilities": other_current_liabilities,
            "Long-term Debt": long_term_debt,
            "Other Non-current Liabilities": other_non_current_liabilities,
            "Equity": equity,
            "Interest Expense": interest_expense,
            "Current Assets": current_assets,
            "Current Liabilities": current_liabilities,
            "Total Assets": total_assets,
            "Total Liabilities": total_liabilities,
            "Retained Earnings": retained_earnings,
            "Market Cap": market_cap,
            "Total Debt": total_debt,
            "Net Debt": net_debt,
            "Enterprise Value": enterprise_value,
            "Invested Capital": invested_capital,
            "Working Capital": working_capital,
        }

    inputs = pd.DataFrame.from_dict(records, orient="index")
    raw_tax_rate = inputs["Taxes"] / inputs["Pretax Income"]
    valid_tax_rate = raw_tax_rate.where(
        inputs["Pretax Income"].gt(0)
        & inputs["Taxes"].ge(0)
        & raw_tax_rate.between(0, 0.60)
    )
    fallback_tax_rate = _to_float(valid_tax_rate.dropna().median())
    inputs["Tax Rate For NOPAT"] = valid_tax_rate.fillna(fallback_tax_rate)
    inputs["NOPAT"] = inputs["EBIT"].where(inputs["EBIT"].le(0), inputs["EBIT"] * (1 - inputs["Tax Rate For NOPAT"]))

    for base, avg_name in [
        ("Total Assets", "Average Assets"),
        ("Equity", "Average Equity"),
        ("Invested Capital", "Average Invested Capital"),
        ("Receivables", "Average Receivables"),
        ("Inventory", "Average Inventory"),
        ("Payables", "Average Payables"),
    ]:
        previous = inputs[base].shift(1)
        averaged = (inputs[base] + previous) / 2
        averaged = averaged.where(inputs[base].notna() & previous.notna(), inputs[base].where(inputs[base].notna(), previous))
        inputs[avg_name] = averaged

    return inputs


def _build_ratio_history(inputs: pd.DataFrame) -> pd.DataFrame:
    ratio_data = pd.DataFrame(index=[definition["label"] for definition in RATIO_DEFINITIONS], columns=inputs.index, dtype=float)

    for year in inputs.index:
        row = {column: _to_float(value) for column, value in inputs.loc[year].items()}

        roic = _safe_div(row["NOPAT"], row["Average Invested Capital"])
        gross_margin = _safe_div(row["Gross Profit"], row["Revenue"])
        operating_margin = _safe_div(row["EBIT"], row["Revenue"])
        net_margin = _safe_div(row["Net Income"], row["Revenue"])
        roe = _safe_div(row["Net Income"], row["Average Equity"])
        roa = _safe_div(row["Net Income"], row["Average Assets"])

        ratio_data.at["ROIC", year] = roic * 100 if roic is not None else None
        ratio_data.at["Gross margin", year] = gross_margin * 100 if gross_margin is not None else None
        ratio_data.at["Operating margin", year] = operating_margin * 100 if operating_margin is not None else None
        ratio_data.at["Net margin", year] = net_margin * 100 if net_margin is not None else None
        ratio_data.at["ROE", year] = roe * 100 if roe is not None else None
        ratio_data.at["ROA", year] = roa * 100 if roa is not None else None

        ratio_data.at["Current ratio", year] = _safe_div(row["Current Assets"], row["Current Liabilities"])
        quick_assets = row["Current Assets"] - row["Inventory"] if None not in (row["Current Assets"], row["Inventory"]) else None
        ratio_data.at["Quick ratio", year] = _safe_div(quick_assets, row["Current Liabilities"])
        ratio_data.at["Cash ratio", year] = _safe_div(row["Cash"], row["Current Liabilities"])
        ratio_data.at["Receivable turnover", year] = _safe_div(row["Revenue"], row["Average Receivables"])

        days_sales = _safe_div(row["Average Receivables"] * 365, row["Revenue"]) if row["Average Receivables"] is not None else None
        days_inventory = _safe_div(row["Average Inventory"] * 365, row["COGS"]) if row["Average Inventory"] is not None else None
        days_payables = _safe_div(row["Average Payables"] * 365, row["COGS"]) if row["Average Payables"] is not None else None
        ratio_data.at["Cash conversion cycle", year] = days_sales + days_inventory - days_payables if None not in (days_sales, days_inventory, days_payables) else None

        altman = None
        if None not in (row["Working Capital"], row["Retained Earnings"], row["EBIT"], row["Market Cap"], row["Total Liabilities"], row["Revenue"], row["Total Assets"]):
            wc_term = _safe_div(row["Working Capital"], row["Total Assets"])
            re_term = _safe_div(row["Retained Earnings"], row["Total Assets"])
            ebit_term = _safe_div(row["EBIT"], row["Total Assets"])
            mve_term = _safe_div(row["Market Cap"], row["Total Liabilities"])
            sales_term = _safe_div(row["Revenue"], row["Total Assets"])
            if None not in (wc_term, re_term, ebit_term, mve_term, sales_term):
                altman = 1.2 * wc_term + 1.4 * re_term + 3.3 * ebit_term + 0.6 * mve_term + 1.0 * sales_term
        ratio_data.at["Altman Z score", year] = altman
        ratio_data.at["Net Debt/EBITDA", year] = _safe_div(row["Net Debt"], row["EBITDA"])
        debt_to_equity = _safe_div(row["Total Debt"], row["Equity"])
        debt_to_capital = _safe_div(row["Total Debt"], (row["Total Debt"] + row["Equity"]) if row["Equity"] is not None else None)
        debt_to_assets = _safe_div(row["Total Debt"], row["Total Assets"])
        financial_leverage = _safe_div(row["Total Liabilities"], row["Total Assets"])
        ratio_data.at["Debt to Equity", year] = debt_to_equity * 100 if debt_to_equity is not None else None
        ratio_data.at["Debt to Capital", year] = debt_to_capital * 100 if debt_to_capital is not None else None
        ratio_data.at["Debt to Assets", year] = debt_to_assets * 100 if debt_to_assets is not None else None
        ratio_data.at["Financial leverage", year] = financial_leverage * 100 if financial_leverage is not None else None
        ratio_data.at["EBITDA to Interest", year] = _safe_div(row["EBITDA"], row["Interest Expense"])

        ratio_data.at["Days of sales", year] = days_sales
        ratio_data.at["Days of inventory", year] = days_inventory
        ratio_data.at["Inventory turnover", year] = _safe_div(row["COGS"], row["Average Inventory"])
        ratio_data.at["Payable turnover", year] = _safe_div(row["COGS"], row["Average Payables"])
        ratio_data.at["Days of payables", year] = days_payables
        ratio_data.at["Asset turnover", year] = _safe_div(row["Revenue"], row["Average Assets"])
        ratio_data.at["Leverage ratio", year] = _safe_div(row["Total Assets"], row["Equity"])

        market_cap = row["Market Cap"]
        enterprise_value = row["Enterprise Value"]
        cfo = row["CFO"]
        net_income = row["Net Income"]
        revenue = row["Revenue"]
        equity = row["Equity"]
        ebitda = row["EBITDA"]

        ratio_data.at["P/E", year] = _safe_div(market_cap, net_income) if None not in (market_cap, net_income) and net_income > 0 else None
        ratio_data.at["P/S", year] = _safe_div(market_cap, revenue) if None not in (market_cap, revenue) and revenue > 0 else None
        ratio_data.at["P/Cash Flow", year] = _safe_div(market_cap, cfo) if None not in (market_cap, cfo) and cfo > 0 else None
        ratio_data.at["Price / Book", year] = _safe_div(market_cap, equity) if None not in (market_cap, equity) and equity > 0 else None
        ratio_data.at["Enterprise Value / EBITDA", year] = (
            _safe_div(enterprise_value, ebitda) if None not in (enterprise_value, ebitda) and enterprise_value > 0 and ebitda > 0 else None
        )
        ratio_data.at["Enterprise Value / Sales", year] = (
            _safe_div(enterprise_value, revenue) if None not in (enterprise_value, revenue) and enterprise_value > 0 and revenue > 0 else None
        )

    return ratio_data


def _build_category_table(category: str, ratio_history: pd.DataFrame) -> tuple[pd.DataFrame, float | None]:
    rows = []
    scores = []

    for definition in [item for item in RATIO_DEFINITIONS if item["category"] == category]:
        series = pd.to_numeric(ratio_history.loc[definition["label"]], errors="coerce")
        score = _score_series(series, definition["kind"], definition.get("target"))
        if score is not None:
            scores.append(float(score))

        if category == "Market multiples":
            clean = pd.to_numeric(series, errors="coerce").dropna()
            latest_value = _to_float(clean.iloc[-1]) if not clean.empty else None
            row = {
                "Ratio": definition["label"],
                "Mark": definition["mark"],
                "Current": _format_value(latest_value),
                "Stars": stars_text(score) if score is not None else "n/a",
                "Score": f"{score}/5" if score is not None else "n/a",
            }
        else:
            years = list(ratio_history.columns)[::-1]
            row = {
                "Ratio": definition["label"],
                "Mark": definition["mark"],
                "Stars": stars_text(score) if score is not None else "n/a",
                "Score": f"{score}/5" if score is not None else "n/a",
            }
            for year in years:
                value = _to_float(series.get(year))
                row[str(year)] = _format_value(value)
            row["Avg"] = _format_value(_to_float(series.mean(skipna=True)))
        rows.append(row)

    category_table = pd.DataFrame(rows)
    category_score = round(sum(scores) / len(scores), 2) if scores else None
    return category_table, category_score


def build_ratio_scorecard(dataset) -> RatioScorecard:
    inputs = _build_inputs(dataset)
    ratio_history = _build_ratio_history(inputs)

    category_tables: dict[str, pd.DataFrame] = {}
    summary_scores: dict[str, float | None] = {}
    notes: list[str] = []

    for category in ["Profitability", "Liquidity", "Credit risk", "Activity"]:
        table, score = _build_category_table(category, ratio_history)
        category_tables[category] = table
        summary_scores[category] = score

    table, score = _build_category_table("Market multiples", ratio_history)
    category_tables["Market multiples"] = table
    summary_scores["Market multiples"] = score

    category_score_values = [score for score in summary_scores.values() if score is not None]
    summary_scores["Total score"] = round(sum(category_score_values) / len(category_score_values), 2) if category_score_values else None

    if ratio_history.loc["Altman Z score"].dropna().empty:
        notes.append("Altman Z score is only shown when Yahoo Finance provides retained earnings, total assets, total liabilities, and market capitalization for the same fiscal year.")
    if ratio_history.loc["P/E"].dropna().empty:
        notes.append(
            "Market-multiple stars use explicit valuation ceilings and the current market value versus the latest annual fundamentals. If earnings, cash flow, or equity are non-positive, those rows stay `n/a`."
        )

    derived_rows = [
        "Working Capital",
        "Retained Earnings",
        "Market Cap",
        "Enterprise Value",
        "Total Debt",
        "Net Debt",
        "Invested Capital",
        "Average Invested Capital",
        "Tax Rate For NOPAT",
        "NOPAT",
        "CFO",
        "Current Assets",
        "Current Liabilities",
        "Total Assets",
        "Total Liabilities",
        "Average Assets",
        "Average Equity",
        "Average Receivables",
        "Average Inventory",
        "Average Payables",
    ]
    derived_table = inputs[derived_rows].T
    derived_table.columns = [str(column) for column in derived_table.columns]
    for column in derived_table.columns:
        derived_table[column] = derived_table[column].map(_format_value)

    formula_table = pd.DataFrame(
        [
            {"Category": definition["category"], "Ratio": definition["label"], "Mark": definition["mark"], "Formula": definition["formula"]}
            for definition in RATIO_DEFINITIONS
        ]
    )

    return RatioScorecard(
        summary_scores=summary_scores,
        category_tables=category_tables,
        derived_table=derived_table,
        formula_table=formula_table,
        notes=notes,
    )
