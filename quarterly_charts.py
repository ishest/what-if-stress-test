from __future__ import annotations

import pandas as pd

from stress_backend import fetch_quarterly_financials, map_historical_financials


_REQUIRED_COLUMNS = [
    "Revenue",
    "Gross Profit",
    "EBIT",
    "Net Income",
    "Total Assets",
    "Current Assets",
    "Total Liabilities",
    "Current Liabilities",
    "Cash & Equivalents",
    "Short-term Debt",
    "Long-term Debt",
    "Equity",
    "Pretax Income",
    "Taxes",
]


def _ensure_numeric_columns(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for column in columns:
        if column not in frame.columns:
            frame[column] = float("nan")
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    return frame


def _average_balance(series: pd.Series) -> pd.Series:
    annual_avg = (series + series.shift(4)) / 2
    sequential_avg = (series + series.shift(1)) / 2
    return annual_avg.where(series.shift(4).notna(), sequential_avg)


def build_quarterly_chart_frame(symbol: str) -> tuple[pd.DataFrame, list[str]]:
    quarterly_raw = fetch_quarterly_financials(symbol.upper().strip())
    quarterly_financials, _, quarterly_warnings = map_historical_financials(
        quarterly_raw,
        label_column="Period Label",
    )

    if (
        quarterly_raw is None
        or quarterly_financials is None
        or quarterly_raw.empty
        or quarterly_financials.empty
        or "Period Label" not in quarterly_raw.columns
        or "asOfDate" not in quarterly_raw.columns
    ):
        return pd.DataFrame(), sorted(set(quarterly_warnings))

    quarterly_dates = quarterly_raw[["Period Label", "asOfDate"]].copy()
    quarterly_dates["asOfDate"] = pd.to_datetime(quarterly_dates["asOfDate"], errors="coerce")
    quarterly_dates = quarterly_dates.dropna(subset=["Period Label", "asOfDate"]).drop_duplicates(subset=["Period Label"])
    if quarterly_dates.empty:
        return pd.DataFrame(), sorted(set(quarterly_warnings))

    quarterly_wide = quarterly_financials.T.reset_index().rename(columns={"index": "Period Label"})
    if quarterly_wide.empty or "Period Label" not in quarterly_wide.columns:
        return pd.DataFrame(), sorted(set(quarterly_warnings))

    frame = quarterly_dates.merge(quarterly_wide, on="Period Label", how="inner").sort_values("asOfDate").reset_index(drop=True)
    if frame.empty:
        return frame, sorted(set(quarterly_warnings))

    frame["Quarter Label"] = frame["asOfDate"].dt.strftime("%b %Y")
    frame = _ensure_numeric_columns(frame, _REQUIRED_COLUMNS)

    frame["Non-current Assets"] = frame["Total Assets"] - frame["Current Assets"]
    frame["Non-current Liabilities"] = frame["Total Liabilities"] - frame["Current Liabilities"]
    frame["Total Debt"] = frame["Short-term Debt"].fillna(0.0) + frame["Long-term Debt"].fillna(0.0)
    frame["Invested Capital"] = frame["Total Debt"] + frame["Equity"] - frame["Cash & Equivalents"].fillna(0.0)

    revenue = frame["Revenue"].replace(0.0, pd.NA)
    frame["Gross Margin %"] = frame["Gross Profit"] / revenue
    frame["EBIT Margin %"] = frame["EBIT"] / revenue
    frame["Net Margin %"] = frame["Net Income"] / revenue

    net_income_ttm = frame["Net Income"].rolling(4, min_periods=4).sum()
    ebit_ttm = frame["EBIT"].rolling(4, min_periods=4).sum()
    pretax_ttm = frame["Pretax Income"].rolling(4, min_periods=4).sum()
    taxes_ttm = frame["Taxes"].rolling(4, min_periods=4).sum()

    tax_rate = (taxes_ttm / pretax_ttm.replace(0.0, pd.NA)).where(pretax_ttm > 0)
    tax_rate = tax_rate.clip(lower=0.0, upper=0.35).fillna(0.21)
    nopat_ttm = ebit_ttm * (1 - tax_rate)

    avg_assets = _average_balance(frame["Total Assets"]).replace(0.0, pd.NA)
    avg_equity = _average_balance(frame["Equity"]).replace(0.0, pd.NA)
    avg_invested_capital = _average_balance(frame["Invested Capital"]).replace(0.0, pd.NA)

    frame["ROA TTM %"] = net_income_ttm / avg_assets
    frame["ROE TTM %"] = net_income_ttm / avg_equity
    frame["ROIC TTM %"] = nopat_ttm / avg_invested_capital
    return frame, sorted(set(quarterly_warnings))
