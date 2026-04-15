from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable

import pandas as pd
import yfinance as yf


MILLION = 1_000_000.0


@dataclass
class QuarterlyChartBundle:
    symbol: str
    revenue_profit: pd.DataFrame
    assets_liabilities: pd.DataFrame
    profitability: pd.DataFrame
    margins: pd.DataFrame
    warnings: list[str]


def _normalize_label(label: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", str(label or ""))


def _coalesce_duplicate_columns(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or frame.columns.is_unique:
        return frame

    combined: dict[str, pd.Series] = {}
    for column in dict.fromkeys(map(str, frame.columns)):
        duplicate_frame = frame.loc[:, frame.columns == column]
        if isinstance(duplicate_frame, pd.Series):
            combined[column] = duplicate_frame
        else:
            combined[column] = duplicate_frame.bfill(axis=1).iloc[:, 0]
    return pd.DataFrame(combined, index=frame.index)


def _prepare_statement_frame(statement: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(statement, pd.DataFrame) or statement.empty:
        return pd.DataFrame()

    frame = statement.T.copy()
    frame.index = pd.to_datetime(frame.index, errors="coerce")
    frame = frame[~frame.index.isna()]
    if frame.empty:
        return pd.DataFrame()

    frame.columns = [_normalize_label(column) for column in frame.columns]
    frame = _coalesce_duplicate_columns(frame)
    return frame.sort_index()


def _get_statement_series(frame: pd.DataFrame, candidates: Iterable[str]) -> pd.Series:
    if frame.empty:
        return pd.Series(dtype="float64")

    result = pd.Series(index=frame.index, dtype="float64")
    found = False
    for candidate in candidates:
        if candidate not in frame.columns:
            continue
        series = pd.to_numeric(frame[candidate], errors="coerce")
        result = result.combine_first(series)
        found = True
    return result if found else pd.Series(index=frame.index, dtype="float64")


def _ttm(series: pd.Series) -> pd.Series:
    if series.empty:
        return series
    return series.rolling(4, min_periods=4).sum()


def _average_balance(series: pd.Series) -> pd.Series:
    if series.empty:
        return series
    return series.rolling(4, min_periods=2).mean()


def _quarter_label(index_value: pd.Timestamp) -> str:
    quarter = index_value.to_period("Q")
    return f"{quarter.year}-Q{quarter.quarter}"


def _finalize_chart_table(frame: pd.DataFrame, *, percent: bool = False) -> pd.DataFrame:
    if frame.empty:
        return frame

    cleaned = frame.copy().sort_index()
    cleaned = cleaned.dropna(how="all")
    if cleaned.empty:
        return cleaned

    if percent:
        cleaned = cleaned * 100.0
    else:
        cleaned = cleaned / MILLION

    cleaned = cleaned.replace([pd.NA, pd.NaT], None)
    cleaned.insert(0, "Quarter", [_quarter_label(value) for value in cleaned.index])
    return cleaned.reset_index(drop=True)


def build_quarterly_chart_bundle(symbol: str, max_periods: int = 12) -> QuarterlyChartBundle:
    symbol = symbol.upper().strip()
    ticker = yf.Ticker(symbol)
    warnings: list[str] = []

    frames = []
    for attr_name, label in [
        ("quarterly_income_stmt", "income statement"),
        ("quarterly_balance_sheet", "balance sheet"),
        ("quarterly_cashflow", "cash flow"),
    ]:
        try:
            prepared = _prepare_statement_frame(getattr(ticker, attr_name))
        except Exception as exc:
            prepared = pd.DataFrame()
            warnings.append(f"Quarterly {label} fetch failed: {exc}")
        if not prepared.empty:
            frames.append(prepared)

    if not frames:
        raise ValueError(f"Yahoo Finance did not return usable quarterly statements for ticker '{symbol}'.")

    combined = pd.concat(frames, axis=1)
    combined = _coalesce_duplicate_columns(combined)
    combined = combined.sort_index()
    if max_periods > 0:
        combined = combined.tail(max_periods)

    revenue = _get_statement_series(combined, ["TotalRevenue", "OperatingRevenue"])
    gross_profit = _get_statement_series(combined, ["GrossProfit"])
    cogs = _get_statement_series(combined, ["CostOfRevenue", "ReconciledCostOfRevenue"]).abs()
    ebit = _get_statement_series(combined, ["EBIT", "TotalOperatingIncomeAsReported"])
    pretax_income = _get_statement_series(combined, ["PretaxIncome"])
    tax_provision = _get_statement_series(combined, ["TaxProvision"]).abs()
    net_income = _get_statement_series(
        combined,
        [
            "NetIncome",
            "NetIncomeCommonStockholders",
            "NetIncomeFromContinuingOperationNetMinorityInterest",
        ],
    )

    if gross_profit.isna().all() and not revenue.isna().all() and not cogs.isna().all():
        gross_profit = revenue - cogs

    cash = _get_statement_series(
        combined,
        ["CashAndCashEquivalents", "CashCashEquivalentsAndShortTermInvestments", "EndCashPosition"],
    )
    receivables = _get_statement_series(combined, ["AccountsReceivable"])
    inventory = _get_statement_series(combined, ["Inventory"])
    other_current_assets = _get_statement_series(combined, ["OtherCurrentAssets"])
    current_assets = _get_statement_series(combined, ["CurrentAssets"])
    total_assets = _get_statement_series(combined, ["TotalAssets"])

    current_debt = _get_statement_series(
        combined,
        ["CurrentDebt", "CurrentDebtAndCapitalLeaseObligation", "OtherCurrentBorrowings"],
    ).fillna(0.0)
    accounts_payable = _get_statement_series(combined, ["AccountsPayable", "Payables"]).fillna(0.0)
    other_current_liabilities = _get_statement_series(combined, ["OtherCurrentLiabilities"]).fillna(0.0)
    current_liabilities = _get_statement_series(combined, ["CurrentLiabilities"])
    long_term_debt = _get_statement_series(
        combined,
        ["LongTermDebt", "LongTermDebtAndCapitalLeaseObligation"],
    ).fillna(0.0)
    total_debt = _get_statement_series(combined, ["TotalDebt"])
    total_liabilities = _get_statement_series(combined, ["TotalLiabilitiesNetMinorityInterest", "TotalLiabilities"])
    equity = _get_statement_series(
        combined,
        ["CommonStockEquity", "StockholdersEquity", "TotalEquityGrossMinorityInterest"],
    )

    if current_assets.isna().all():
        current_assets = cash.fillna(0.0) + receivables.fillna(0.0) + inventory.fillna(0.0) + other_current_assets.fillna(0.0)
    if current_liabilities.isna().all():
        current_liabilities = current_debt + accounts_payable + other_current_liabilities
    if total_liabilities.isna().all() and not total_assets.isna().all() and not equity.isna().all():
        total_liabilities = total_assets - equity

    non_current_assets = total_assets - current_assets
    non_current_liabilities = total_liabilities - current_liabilities

    if total_debt.isna().all():
        total_debt = current_debt + long_term_debt
    total_debt = total_debt.fillna(current_debt + long_term_debt)

    revenue_ttm = _ttm(revenue)
    gross_profit_ttm = _ttm(gross_profit)
    ebit_ttm = _ttm(ebit)
    net_income_ttm = _ttm(net_income)

    pretax_income_ttm = _ttm(pretax_income)
    tax_provision_ttm = _ttm(tax_provision)
    tax_rate = (tax_provision_ttm / pretax_income_ttm.abs()).clip(lower=0.0, upper=0.35)
    tax_rate = tax_rate.fillna(0.21)
    nopat_ttm = ebit_ttm * (1 - tax_rate)

    avg_equity = _average_balance(equity)
    avg_assets = _average_balance(total_assets)
    invested_capital = equity + total_debt - cash.fillna(0.0)
    avg_invested_capital = _average_balance(invested_capital)

    revenue_profit = _finalize_chart_table(
        pd.DataFrame(
            {
                "Revenue": revenue_ttm,
                "Gross Profit": gross_profit_ttm,
                "EBIT": ebit_ttm,
                "Net Income": net_income_ttm,
            }
        )
    )

    assets_liabilities = _finalize_chart_table(
        pd.DataFrame(
            {
                "Total Assets": total_assets,
                "Cash": cash,
                "Current Assets": current_assets,
                "Non-current Assets": non_current_assets,
                "Total Liabilities": total_liabilities,
                "Current Liabilities": current_liabilities,
                "Non-current Liabilities": non_current_liabilities,
            }
        )
    )

    profitability = _finalize_chart_table(
        pd.DataFrame(
            {
                "ROIC": nopat_ttm / avg_invested_capital,
                "ROE": net_income_ttm / avg_equity,
                "ROA": net_income_ttm / avg_assets,
            }
        ),
        percent=True,
    )

    margins = _finalize_chart_table(
        pd.DataFrame(
            {
                "Gross Profit %": gross_profit_ttm / revenue_ttm,
                "EBIT Margin %": ebit_ttm / revenue_ttm,
                "Net Income %": net_income_ttm / revenue_ttm,
            }
        ),
        percent=True,
    )

    if revenue_profit.empty and assets_liabilities.empty and profitability.empty and margins.empty:
        raise ValueError(f"Yahoo Finance quarterly statements for '{symbol}' did not contain enough data to chart.")

    return QuarterlyChartBundle(
        symbol=symbol,
        revenue_profit=revenue_profit,
        assets_liabilities=assets_liabilities,
        profitability=profitability,
        margins=margins,
        warnings=warnings,
    )
