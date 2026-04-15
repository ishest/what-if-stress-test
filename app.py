from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from financial_ratios import CATEGORY_ORDER, build_ratio_scorecard, stars_text
from multiples import build_multiples_snapshot
from quarterly_charts import build_quarterly_chart_frame
from stock_scoring import build_stock_scoring_model
from stress_backend import (
    CompanyOverview,
    DISTRESS_MAX_NET_LEVERAGE,
    DISTRESS_MIN_CURRENT_RATIO,
    DISTRESS_MIN_ENDING_CASH,
    DISTRESS_MIN_INTEREST_COVERAGE,
    StressModelDataError,
    ThresholdSettings,
    build_historical_dataset,
    fetch_company_overview,
    get_selected_scenario,
    load_workbook_model,
    prepare_latest_for_stress,
    run_all_scenarios,
    run_scenario,
)


st.set_page_config(
    page_title="What If Stress Test",
    page_icon="🚦",
    layout="wide",
    initial_sidebar_state="expanded",
)


SEVERITY_ORDER = {"Light": 1, "Base": 2, "Severe": 3}
RATING_ORDER = {"CRITICAL": 1, "WATCH": 2, "RESILIENT": 3}
SEVERITY_BG = {"Light": "#d9efe1", "Base": "#f8e8c8", "Severe": "#f6d7d7"}
RATING_BG = {"CRITICAL": "#c53d2d", "WATCH": "#f2a31b", "RESILIENT": "#23863a"}
HEATMAP_OVERALL_KEY = "__overall__"
HEATMAP_OVERALL_LABEL = "Overall"
HEATMAP_METRIC_SPECS = [
    ("EBIT / interest", "EBIT / int"),
    ("Net debt / EBITDA", "Net lev."),
    ("Current ratio", "Curr. ratio"),
    ("FCF", "FCF"),
    ("Ending cash", "End cash"),
    ("Ending equity", "End equity"),
]
HEATMAP_SEVERITY_SHORT = {"Light": "L", "Base": "B", "Severe": "S"}
HEATMAP_SEVERITY_WEIGHT = {"Light": 3, "Base": 2, "Severe": 1}
HEATMAP_STATUS_SCORE = {"RESILIENT": 0, "WATCH": 1, "CRITICAL": 2}
HEATMAP_COLORSCALE = [
    (0.00, RATING_BG["RESILIENT"]),
    (0.33, RATING_BG["RESILIENT"]),
    (0.34, RATING_BG["WATCH"]),
    (0.66, RATING_BG["WATCH"]),
    (0.67, RATING_BG["CRITICAL"]),
    (1.00, RATING_BG["CRITICAL"]),
]

HARD_MINIMUM_CASH_BUFFER = 0.0
HARD_MINIMUM_INTEREST_COVERAGE = DISTRESS_MIN_INTEREST_COVERAGE
HARD_MAXIMUM_NET_LEVERAGE = 1.0
HARD_MINIMUM_CURRENT_RATIO = DISTRESS_MIN_CURRENT_RATIO
SGA_SHARE_LOWER_BOUND = 0.05
SGA_SHARE_UPPER_BOUND = 0.80
THRESHOLD_WIDGET_KEYS = {
    "sga_variable_cost_share": "threshold_sga_variable_cost_share",
    "minimum_cash_buffer": "threshold_minimum_cash_buffer",
    "minimum_interest_coverage": "threshold_minimum_interest_coverage",
    "maximum_net_leverage": "threshold_maximum_net_leverage",
    "minimum_current_ratio": "threshold_minimum_current_ratio",
}


def _status_from_score(score: float | int | None) -> str:
    if score is None or pd.isna(score):
        return "RESILIENT"
    if score >= 2:
        return "CRITICAL"
    if score >= 1:
        return "WATCH"
    return "RESILIENT"


@dataclass
class ThresholdCalibration:
    settings: ThresholdSettings
    details: pd.DataFrame
    notes: list[str]


def format_m(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:,.1f}"


def format_signed_m(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:+,.1f}"


def format_pct(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:.1%}"


def format_signed_pct(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:+.1%}"


def format_ratio(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:.2f}x"


def format_signed_ratio(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:+.2f}x"


def format_days(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{value:.1f}d"


def style_rating(value: str) -> str:
    if value == "CRITICAL":
        return "critical"
    if value == "WATCH":
        return "watch"
    return "resilient"


def render_color_card(title: str, value: str, background: str, text_color: str = "#ffffff"):
    st.markdown(
        f"""
        <div style="
            background:{background};
            color:{text_color};
            padding:18px 12px;
            border-radius:14px;
            text-align:center;
            font-weight:700;
            min-height:86px;
            display:flex;
            flex-direction:column;
            justify-content:center;
            box-shadow:0 1px 2px rgba(0,0,0,0.08);
        ">
            <div style="font-size:15px; opacity:0.95;">{title}</div>
            <div style="font-size:32px; line-height:1.1; margin-top:4px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_star_card(title: str, score: float | None):
    score_label = "n/a" if score is None or pd.isna(score) else f"{score:.1f}/5"
    stars = stars_text(score)
    st.markdown(
        f"""
        <div style="
            background:#ffffff;
            color:#1b2a41;
            padding:16px 12px;
            border-radius:14px;
            border:1px solid #d9dde5;
            min-height:110px;
            display:flex;
            flex-direction:column;
            justify-content:center;
            box-shadow:0 1px 2px rgba(0,0,0,0.05);
        ">
            <div style="font-size:15px; font-weight:700; text-align:center;">{title}</div>
            <div style="font-size:28px; text-align:center; letter-spacing:1px; margin-top:8px;">{stars}</div>
            <div style="font-size:14px; color:#5f6b7a; text-align:center; margin-top:6px;">{score_label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_message_card(title: str, body: str, background: str, text_color: str = "#1b2a41"):
    st.markdown(
        f"""
        <div style="
            background:{background};
            color:{text_color};
            padding:16px 16px;
            border-radius:14px;
            min-height:150px;
            border:1px solid rgba(27,42,65,0.08);
            box-shadow:0 1px 2px rgba(0,0,0,0.04);
        ">
            <div style="font-size:15px; font-weight:700; margin-bottom:8px;">{title}</div>
            <div style="font-size:14px; line-height:1.5; white-space:pre-wrap;">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def get_dashboard_matrix(scenario_matrix: pd.DataFrame, scenario_library: pd.DataFrame) -> pd.DataFrame:
    sequence_order_map = {}
    order = 1
    for sequence in scenario_library["Sequence"]:
        if pd.isna(sequence) or sequence == "Custom / Blank" or sequence in sequence_order_map:
            continue
        sequence_order_map[sequence] = order
        order += 1

    dashboard_matrix = scenario_matrix[scenario_matrix["Sequence"] != "Custom / Blank"].copy()
    dashboard_matrix["Sequence Order"] = dashboard_matrix["Sequence"].map(sequence_order_map).fillna(999)
    dashboard_matrix["Severity Order"] = dashboard_matrix["Severity"].map(SEVERITY_ORDER).fillna(999)
    dashboard_matrix = dashboard_matrix.sort_values(["Sequence Order", "Severity Order"]).reset_index(drop=True)
    return dashboard_matrix


def base_current_ratio(latest_values: dict[str, float]) -> float:
    return (
        latest_values["Cash & Equivalents"]
        + latest_values["Accounts Receivable"]
        + latest_values["Inventory"]
        + latest_values.get("Other Current Assets", 0.0)
    ) / max(
        latest_values.get("Accounts Payable", 0.0)
        + latest_values.get("Other Current Liabilities", 0.0)
        + latest_values.get("Short-term Debt", 0.0),
        0.01,
    )


def build_critical_points(dataset, dashboard_matrix: pd.DataFrame, thresholds: ThresholdSettings) -> dict[str, pd.DataFrame | list[str] | dict[str, float]]:
    pressure_counts: dict[str, int] = {}
    critical_counts: dict[str, int] = {}
    for status_map in dashboard_matrix["Dashboard Status"]:
        for metric, status in status_map.items():
            if status in {"WATCH", "CRITICAL"}:
                pressure_counts[metric] = pressure_counts.get(metric, 0) + 1
            if status == "CRITICAL":
                critical_counts[metric] = critical_counts.get(metric, 0) + 1

    failure_rows = []
    total_scenarios = len(dashboard_matrix)
    for metric, count in sorted(pressure_counts.items(), key=lambda item: (-item[1], item[0])):
        failure_rows.append(
            {
                "Failure point": metric,
                "Fail count": count,
                "Critical count": critical_counts.get(metric, 0),
                "Share of scenarios": count / total_scenarios if total_scenarios else 0.0,
            }
        )
    failure_df = pd.DataFrame(failure_rows)

    breakpoint_rows = []
    for sequence, group in dashboard_matrix.groupby("Sequence", sort=False):
        ordered = group.sort_values("Severity Order")
        first_watch = ordered[ordered["Rating"].isin(["WATCH", "CRITICAL"])].head(1)
        first_critical = ordered[ordered["Rating"] == "CRITICAL"].head(1)
        rated = ordered.assign(_rating_order=ordered["Rating"].map(RATING_ORDER).fillna(99))
        worst_row = rated.sort_values(["_rating_order", "Severity Order"]).iloc[0]
        breakpoint_rows.append(
            {
                "Sequence": sequence,
                "First non-resilient severity": first_watch.iloc[0]["Severity"] if not first_watch.empty else "None",
                "First critical severity": first_critical.iloc[0]["Severity"] if not first_critical.empty else "None",
                "Worst rating": worst_row["Rating"],
            }
        )
    breakpoints_df = pd.DataFrame(breakpoint_rows)

    worst_cases = dashboard_matrix.copy()
    worst_cases["Rating Order"] = worst_cases["Rating"].map(RATING_ORDER).fillna(99)
    worst_cases = worst_cases.sort_values(
        ["Rating Order", "Ending Cash", "Interest Coverage", "Net Leverage", "FCF (Stressed)"],
        ascending=[True, True, True, False, True],
    ).head(10)
    worst_cases = worst_cases[
        [
            "Sequence",
            "Severity",
            "Rating",
            "Rating Reasons",
            "Ending Cash",
            "FCF (Stressed)",
            "Net Leverage",
            "Interest Coverage",
            "Ending Equity",
        ]
    ].copy()

    latest = dataset.latest_values
    base_current = base_current_ratio(latest)
    headroom = {
        "Net leverage benchmark headroom": thresholds.maximum_net_leverage - latest["Net Leverage"],
        "Net leverage distress headroom": DISTRESS_MAX_NET_LEVERAGE - latest["Net Leverage"],
        "Interest coverage benchmark headroom": latest["Interest Coverage"] - thresholds.minimum_interest_coverage,
        "Interest coverage distress headroom": latest["Interest Coverage"] - DISTRESS_MIN_INTEREST_COVERAGE,
        "Current ratio benchmark headroom": base_current - thresholds.minimum_current_ratio,
        "Current ratio distress headroom": base_current - DISTRESS_MIN_CURRENT_RATIO,
        "Cash vs benchmark buffer": latest["Cash & Equivalents"] - thresholds.minimum_cash_buffer,
        "Cash vs zero": latest["Cash & Equivalents"] - DISTRESS_MIN_ENDING_CASH,
    }

    insights: list[str] = []
    if not failure_df.empty:
        top_failure = failure_df.iloc[0]
        insights.append(
            f"The most common pressure point is `{top_failure['Failure point']}`, which turns non-OK in {int(top_failure['Fail count'])} of {total_scenarios} scenarios."
        )

    early_breaks = breakpoints_df[breakpoints_df["First non-resilient severity"].isin(["Light", "Base"])]
    if not early_breaks.empty:
        early_list = ", ".join(
            f"{row['Sequence']} ({row['First non-resilient severity']})" for _, row in early_breaks.head(4).iterrows()
        )
        insights.append(f"The earliest fragility shows up in: {early_list}.")

    if headroom["Net leverage distress headroom"] < 0:
        insights.append("Base leverage is already above the hard distress ceiling before stress is applied.")
    elif headroom["Net leverage benchmark headroom"] < 0:
        insights.append(
            "Base leverage is above the company benchmark, but still below the hard distress ceiling."
        )
    if headroom["Interest coverage distress headroom"] < 0:
        insights.append("Base interest coverage is already below the hard distress floor before stress is applied.")
    elif headroom["Interest coverage benchmark headroom"] < 0:
        insights.append(
            "Base interest coverage is below the company benchmark, but still above the hard distress floor."
        )
    if headroom["Current ratio distress headroom"] < 0:
        insights.append("Base current ratio is already below the hard distress floor before stress is applied.")
    elif headroom["Current ratio benchmark headroom"] < 0:
        insights.append("Base current ratio is below the company benchmark, but still above the hard distress floor.")
    if latest["FCF (CFO - Capex)"] < 0:
        insights.append("Base free cash flow is already negative before stress, which makes liquidity-driven scenarios more dangerous.")
    if headroom["Cash vs zero"] < 0:
        insights.append("Base cash is already negative, which is a clear liquidity stress signal.")
    elif headroom["Cash vs benchmark buffer"] < 0:
        insights.append("Base cash is below the selected company buffer, but still above zero.")

    heatmap_rows: list[dict[str, str | int]] = []

    def _normalize_metric_status(status: str) -> str:
        if status == "CRITICAL":
            return "CRITICAL"
        if status in {"WATCH", "INFO"}:
            return "WATCH"
        return "RESILIENT"

    for sequence, group in dashboard_matrix.groupby("Sequence", sort=False):
        ordered = group.sort_values("Severity Order")
        sequence_order = int(ordered["Sequence Order"].iloc[0]) if not ordered.empty else 999

        for _, row in ordered.iterrows():
            severity = row["Severity"]
            severity_order = int(row["Severity Order"])

            heatmap_rows.append(
                {
                    "Sequence": sequence,
                    "Sequence Order": sequence_order,
                    "Severity": severity,
                    "Severity Order": severity_order,
                    "Severity Short": HEATMAP_SEVERITY_SHORT.get(severity, severity[:1]),
                    "Metric Key": HEATMAP_OVERALL_KEY,
                    "Metric Label": HEATMAP_OVERALL_LABEL,
                    "Status": row["Rating"],
                    "Status Score": HEATMAP_STATUS_SCORE.get(row["Rating"], 0),
                }
            )

            for metric_key, metric_label in HEATMAP_METRIC_SPECS:
                metric_status = _normalize_metric_status(row["Dashboard Status"].get(metric_key, "OK"))
                heatmap_rows.append(
                    {
                        "Sequence": sequence,
                        "Sequence Order": sequence_order,
                        "Severity": severity,
                        "Severity Order": severity_order,
                        "Severity Short": HEATMAP_SEVERITY_SHORT.get(severity, severity[:1]),
                        "Metric Key": metric_key,
                        "Metric Label": metric_label,
                        "Status": metric_status,
                        "Status Score": HEATMAP_STATUS_SCORE.get(metric_status, 0),
                    }
                )

    heatmap_df = pd.DataFrame(heatmap_rows)

    return {
        "failure_df": failure_df,
        "breakpoints_df": breakpoints_df,
        "worst_cases": worst_cases,
        "headroom": headroom,
        "insights": insights,
        "heatmap_df": heatmap_df,
    }


def render_failure_heatmap(heatmap_df: pd.DataFrame):
    st.markdown("**Failure Heatmap**")
    control_col1, control_col2 = st.columns([1.4, 1.2])
    with control_col1:
        heatmap_view = st.segmented_control(
            "View",
            ["Worst case view", "By severity"],
            default="Worst case view",
            key="critical_points_heatmap_view",
            width="stretch",
        )
    with control_col2:
        heatmap_sort = st.segmented_control(
            "Sort rows",
            ["Fragility", "Scenario order"],
            default="Fragility",
            key="critical_points_heatmap_sort",
            width="stretch",
        )

    st.caption(
        "Click any cell to preload the Scenario Explorer. Green means resilient, amber means pressure, and red means the metric reaches a hard distress zone."
    )

    if heatmap_df.empty:
        st.info("No heatmap is available because there are no scenario results to summarize.")
        return

    non_overall = heatmap_df[heatmap_df["Metric Key"] != HEATMAP_OVERALL_KEY].copy()
    summary_rows = []
    for sequence, group in heatmap_df.groupby("Sequence", sort=False):
        non_overall_group = group[group["Metric Key"] != HEATMAP_OVERALL_KEY].copy()
        overall_group = group[group["Metric Key"] == HEATMAP_OVERALL_KEY].copy()
        summary_rows.append(
            {
                "Sequence": sequence,
                "Sequence Order": int(group["Sequence Order"].iloc[0]),
                "Critical Cells": int((non_overall_group["Status"] == "CRITICAL").sum()),
                "Pressure Cells": int(non_overall_group["Status"].isin(["WATCH", "CRITICAL"]).sum()),
                "Fragility Score": float(
                    (
                        non_overall_group["Status Score"]
                        * non_overall_group["Severity"].map(HEATMAP_SEVERITY_WEIGHT).fillna(1)
                    ).sum()
                ),
                "Worst Overall Score": int(overall_group["Status Score"].max()) if not overall_group.empty else 0,
                "First Overall Pressure Order": int(overall_group.loc[overall_group["Status"].isin(["WATCH", "CRITICAL"]), "Severity Order"].min()) if overall_group["Status"].isin(["WATCH", "CRITICAL"]).any() else 99,
                "First Overall Critical Order": int(overall_group.loc[overall_group["Status"] == "CRITICAL", "Severity Order"].min()) if (overall_group["Status"] == "CRITICAL").any() else 99,
            }
        )
    summary_df = pd.DataFrame(summary_rows)
    if heatmap_sort == "Fragility":
        sequence_order = summary_df.sort_values(
            [
                "Worst Overall Score",
                "First Overall Critical Order",
                "Critical Cells",
                "Pressure Cells",
                "Fragility Score",
                "Sequence Order",
            ],
            ascending=[False, True, False, False, False, True],
        )["Sequence"].tolist()
    else:
        sequence_order = summary_df.sort_values("Sequence Order")["Sequence"].tolist()

    overall_summary_rows = []
    overall_detail = heatmap_df[heatmap_df["Metric Key"] == HEATMAP_OVERALL_KEY].copy()
    for sequence, group in overall_detail.groupby("Sequence", sort=False):
        ordered = group.sort_values("Severity Order")
        scores = ordered["Status Score"].tolist()
        worst_score = max(scores) if scores else 0
        worst_status = _status_from_score(worst_score)
        first_non_resilient = ordered.loc[ordered["Status"].isin(["WATCH", "CRITICAL"]), "Severity"].iloc[0] if ordered["Status"].isin(["WATCH", "CRITICAL"]).any() else "None"
        first_critical = ordered.loc[ordered["Status"] == "CRITICAL", "Severity"].iloc[0] if (ordered["Status"] == "CRITICAL").any() else "None"
        target_severity = first_critical if first_critical != "None" else first_non_resilient if first_non_resilient != "None" else "Light"
        overall_summary_rows.append(
            {
                "Sequence": sequence,
                "Column Label": HEATMAP_OVERALL_LABEL,
                "Status": worst_status,
                "Status Score": worst_score,
                "Target Severity": target_severity,
                "Metric Key": HEATMAP_OVERALL_KEY,
                "Metric Label": HEATMAP_OVERALL_LABEL,
                "Trigger Detail": (
                    "Remains resilient in Light, Base, and Severe."
                    if worst_status == "RESILIENT"
                    else f"First non-resilient at {first_non_resilient}."
                    if worst_status == "WATCH"
                    else f"First non-resilient at {first_non_resilient}; first critical at {first_critical}."
                ),
                "Severity Path": " | ".join(f"{row['Severity']}: {row['Status']}" for _, row in ordered.iterrows()),
            }
        )
    overall_summary_df = pd.DataFrame(overall_summary_rows)

    if heatmap_view == "Worst case view":
        display_rows = []
        for sequence, group in non_overall.groupby("Sequence", sort=False):
            for metric_key, metric_label in HEATMAP_METRIC_SPECS:
                metric_group = group[group["Metric Key"] == metric_key].sort_values("Severity Order")
                scores = metric_group["Status Score"].tolist()
                worst_score = max(scores) if scores else 0
                worst_status = _status_from_score(worst_score)
                first_non_resilient = metric_group.loc[metric_group["Status"].isin(["WATCH", "CRITICAL"]), "Severity"].iloc[0] if metric_group["Status"].isin(["WATCH", "CRITICAL"]).any() else "None"
                first_critical = metric_group.loc[metric_group["Status"] == "CRITICAL", "Severity"].iloc[0] if (metric_group["Status"] == "CRITICAL").any() else "None"
                target_severity = first_critical if first_critical != "None" else first_non_resilient if first_non_resilient != "None" else "Light"
                display_rows.append(
                    {
                        "Sequence": sequence,
                        "Column Label": metric_label,
                        "Status": worst_status,
                        "Status Score": worst_score,
                        "Target Severity": target_severity,
                        "Metric Key": metric_key,
                        "Metric Label": metric_label,
                        "Trigger Detail": (
                            "Remains resilient in Light, Base, and Severe."
                            if worst_status == "RESILIENT"
                            else f"First non-resilient at {first_non_resilient}."
                            if worst_status == "WATCH"
                            else f"First non-resilient at {first_non_resilient}; first critical at {first_critical}."
                        ),
                        "Severity Path": " | ".join(f"{row['Severity']}: {row['Status']}" for _, row in metric_group.iterrows()),
                    }
                )
        display_df = pd.DataFrame(display_rows)
        if not overall_summary_df.empty:
            display_df = pd.concat([display_df, overall_summary_df], ignore_index=True)
        column_order = [label for _, label in HEATMAP_METRIC_SPECS] + [HEATMAP_OVERALL_LABEL]
    else:
        severity_display = non_overall.copy()
        severity_display["Column Label"] = severity_display.apply(
            lambda row: f"{row['Metric Label']}<br>{row['Severity Short']}",
            axis=1,
        )
        severity_display["Target Severity"] = severity_display["Severity"]
        severity_display["Trigger Detail"] = severity_display.apply(
            lambda row: (
                f"Selected severity {row['Severity']} stays resilient."
                if row["Status"] == "RESILIENT"
                else f"Selected severity {row['Severity']} is in WATCH."
                if row["Status"] == "WATCH"
                else f"Selected severity {row['Severity']} is CRITICAL."
            ),
            axis=1,
        )
        path_lookup = (
            non_overall.sort_values("Severity Order")
            .groupby(["Sequence", "Metric Key"])["Status"]
            .agg(list)
            .to_dict()
        )
        severity_display["Severity Path"] = severity_display.apply(
            lambda row: " | ".join(
                f"{severity}: {status}"
                for severity, status in zip(
                    ["Light", "Base", "Severe"],
                    path_lookup.get((row["Sequence"], row["Metric Key"]), []),
                )
            ),
            axis=1,
        )
        display_df = severity_display[
            [
                "Sequence",
                "Column Label",
                "Status",
                "Status Score",
                "Target Severity",
                "Metric Key",
                "Metric Label",
                "Trigger Detail",
                "Severity Path",
            ]
        ].copy()
        if not overall_summary_df.empty:
            display_df = pd.concat([display_df, overall_summary_df], ignore_index=True)
        column_order = [
            f"{metric_label}<br>{HEATMAP_SEVERITY_SHORT[severity]}"
            for metric_key, metric_label in HEATMAP_METRIC_SPECS
            for severity in ["Light", "Base", "Severe"]
        ] + [HEATMAP_OVERALL_LABEL]

    score_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Status Score")
        .reindex(index=sequence_order, columns=column_order)
        .fillna(0)
    )
    status_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Status")
        .reindex(index=sequence_order, columns=column_order)
        .fillna("RESILIENT")
    )
    trigger_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Trigger Detail")
        .reindex(index=sequence_order, columns=column_order)
        .fillna("")
    )
    path_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Severity Path")
        .reindex(index=sequence_order, columns=column_order)
        .fillna("")
    )
    target_severity_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Target Severity")
        .reindex(index=sequence_order, columns=column_order)
        .fillna("Light")
    )
    metric_key_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Metric Key")
        .reindex(index=sequence_order, columns=column_order)
        .fillna(HEATMAP_OVERALL_KEY)
    )
    metric_label_matrix = (
        display_df.pivot(index="Sequence", columns="Column Label", values="Metric Label")
        .reindex(index=sequence_order, columns=column_order)
        .fillna(HEATMAP_OVERALL_LABEL)
    )

    customdata = []
    for sequence in sequence_order:
        custom_row = []
        for column in column_order:
            custom_row.append(
                [
                    sequence,
                    metric_key_matrix.loc[sequence, column],
                    target_severity_matrix.loc[sequence, column],
                    status_matrix.loc[sequence, column],
                    trigger_matrix.loc[sequence, column],
                    path_matrix.loc[sequence, column],
                    metric_label_matrix.loc[sequence, column],
                ]
            )
        customdata.append(custom_row)

    fig = go.Figure(
        data=go.Heatmap(
            z=score_matrix.values,
            x=column_order,
            y=sequence_order,
            colorscale=HEATMAP_COLORSCALE,
            zmin=0,
            zmax=2,
            customdata=customdata,
            xgap=2,
            ygap=2,
            showscale=False,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Metric: %{customdata[6]}<br>"
                "Status: %{customdata[3]}<br>"
                "%{customdata[4]}<br>"
                "Path: %{customdata[5]}"
                "<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        height=max(520, 30 * len(sequence_order) + 120),
        margin=dict(l=20, r=20, t=20, b=20),
        xaxis_title="Metric",
        yaxis_title="Stress sequence",
        yaxis=dict(autorange="reversed"),
    )
    fig.update_xaxes(side="top", tickangle=0, tickfont=dict(size=11))
    event = st.plotly_chart(
        fig,
        key="critical_points_failure_heatmap",
        on_select="rerun",
        selection_mode="points",
        width="stretch",
    )

    selection_state = event.selection if hasattr(event, "selection") else event.get("selection", {})
    points = selection_state.get("points", []) if isinstance(selection_state, dict) else getattr(selection_state, "points", [])
    if points:
        custom = points[-1].get("customdata", [])
        if len(custom) >= 7:
            st.session_state["scenario_explorer_sequence"] = custom[0]
            st.session_state["scenario_explorer_severity"] = custom[2]
            st.session_state["scenario_explorer_metric"] = custom[1]
            st.session_state["scenario_explorer_metric_label"] = custom[6]
            st.session_state["scenario_explorer_source"] = "failure_heatmap"
            st.session_state["critical_points_heatmap_selection"] = {
                "sequence": custom[0],
                "severity": custom[2],
                "metric_key": custom[1],
                "metric_label": custom[6],
                "status": custom[3],
                "trigger_detail": custom[4],
                "path": custom[5],
            }

    selected_cell = st.session_state.get("critical_points_heatmap_selection")
    if selected_cell:
        st.info(
            f"Selected `{selected_cell['sequence']}` -> `{selected_cell['metric_label']}` as `{selected_cell['status']}`. "
            f"Scenario Explorer is preloaded with `{selected_cell['severity']}` for the same sequence."
        )
        st.caption(f"{selected_cell['trigger_detail']} Path: {selected_cell['path']}")


def build_base_threshold_status(dataset, thresholds: ThresholdSettings) -> tuple[pd.DataFrame, list[str], list[str]]:
    latest = dataset.latest_values
    current_ratio = base_current_ratio(latest)
    rows = []
    benchmark_gaps: list[str] = []
    distress_issues: list[str] = []

    metric_specs = [
        {
            "Metric": "Cash buffer",
            "Base Raw": latest["Cash & Equivalents"],
            "Company Benchmark Raw": thresholds.minimum_cash_buffer,
            "Distress Raw": DISTRESS_MIN_ENDING_CASH,
            "Base": format_m(latest["Cash & Equivalents"]),
            "Company Benchmark": f">= {format_m(thresholds.minimum_cash_buffer)}",
            "Distress Zone": "<= 0.0",
            "kind": "min",
        },
        {
            "Metric": "Net debt / EBITDA",
            "Base Raw": latest["Net Leverage"],
            "Company Benchmark Raw": thresholds.maximum_net_leverage,
            "Distress Raw": DISTRESS_MAX_NET_LEVERAGE,
            "Base": format_ratio(latest["Net Leverage"]),
            "Company Benchmark": f"<= {thresholds.maximum_net_leverage:.2f}x",
            "Distress Zone": f">= {DISTRESS_MAX_NET_LEVERAGE:.2f}x",
            "kind": "max",
        },
        {
            "Metric": "EBIT / interest",
            "Base Raw": latest["Interest Coverage"],
            "Company Benchmark Raw": thresholds.minimum_interest_coverage,
            "Distress Raw": DISTRESS_MIN_INTEREST_COVERAGE,
            "Base": format_ratio(latest["Interest Coverage"]),
            "Company Benchmark": f">= {thresholds.minimum_interest_coverage:.2f}x",
            "Distress Zone": f"<= {DISTRESS_MIN_INTEREST_COVERAGE:.2f}x",
            "kind": "min",
        },
        {
            "Metric": "Current ratio",
            "Base Raw": current_ratio,
            "Company Benchmark Raw": thresholds.minimum_current_ratio,
            "Distress Raw": DISTRESS_MIN_CURRENT_RATIO,
            "Base": format_ratio(current_ratio),
            "Company Benchmark": f">= {thresholds.minimum_current_ratio:.2f}x",
            "Distress Zone": f"<= {DISTRESS_MIN_CURRENT_RATIO:.2f}x",
            "kind": "min",
        },
    ]

    for spec in metric_specs:
        base_raw = spec["Base Raw"]
        benchmark_raw = spec["Company Benchmark Raw"]
        distress_raw = spec["Distress Raw"]
        if spec["kind"] == "min":
            benchmark_ok = base_raw >= benchmark_raw
            distress_ok = base_raw > distress_raw
        else:
            benchmark_ok = base_raw <= benchmark_raw
            distress_ok = base_raw < distress_raw

        if not distress_ok:
            interpretation = "Below hard distress floor / beyond hard distress ceiling."
            distress_issues.append(
                f"{spec['Metric']} {spec['Base']} is already in the distress zone {spec['Distress Zone']}."
            )
        elif not benchmark_ok:
            interpretation = "Below the company benchmark, but still within normal financial safety limits."
            benchmark_gaps.append(
                f"{spec['Metric']} {spec['Base']} is below the company benchmark {spec['Company Benchmark']}, but still outside the distress zone {spec['Distress Zone']}."
            )
        else:
            interpretation = "Above the company benchmark and outside the distress zone."

        rows.append(
            {
                "Metric": spec["Metric"],
                "Base": spec["Base"],
                "Company Benchmark": spec["Company Benchmark"],
                "Distress Zone": spec["Distress Zone"],
                "Benchmark Status": "OK" if benchmark_ok else "WATCH",
                "Distress Status": "SAFE" if distress_ok else "CRITICAL",
                "Interpretation": interpretation,
            }
        )

    base_status = pd.DataFrame(rows)
    return base_status, benchmark_gaps, distress_issues


def build_result_explanation(dataset, dashboard_matrix: pd.DataFrame, thresholds: ThresholdSettings) -> dict[str, Any]:
    critical_points = build_critical_points(dataset, dashboard_matrix, thresholds)
    base_status, base_benchmark_gaps, base_distress_issues = build_base_threshold_status(dataset, thresholds)

    total_scenarios = len(dashboard_matrix)
    critical_count = int((dashboard_matrix["Rating"] == "CRITICAL").sum())
    watch_count = int((dashboard_matrix["Rating"] == "WATCH").sum())
    critical_by_severity = (
        dashboard_matrix[dashboard_matrix["Rating"] == "CRITICAL"]["Severity"].value_counts().to_dict()
    )
    light_critical = int(critical_by_severity.get("Light", 0))
    base_critical = int(critical_by_severity.get("Base", 0))
    severe_critical = int(critical_by_severity.get("Severe", 0))

    def _count_flags(series: pd.Series) -> dict[str, int]:
        counts: dict[str, int] = {}
        for flag_map in series:
            if not isinstance(flag_map, dict):
                continue
            for label, is_triggered in flag_map.items():
                if is_triggered:
                    counts[label] = counts.get(label, 0) + 1
        return counts

    critical_reason_counts = _count_flags(
        dashboard_matrix.loc[dashboard_matrix["Rating"] == "CRITICAL", "Critical Flags"]
    )
    benchmark_reason_counts = _count_flags(
        dashboard_matrix.loc[dashboard_matrix["Rating"].isin(["WATCH", "CRITICAL"]), "Benchmark Flags"]
    )

    reason_rows = []
    for label, count in critical_reason_counts.items():
        reason_rows.append(
            {
                "Type": "Critical trigger",
                "Driver": label,
                "Count": count,
                "Share of scenarios": count / total_scenarios if total_scenarios else 0.0,
            }
        )
    for label, count in benchmark_reason_counts.items():
        reason_rows.append(
            {
                "Type": "Benchmark pressure",
                "Driver": label,
                "Count": count,
                "Share of scenarios": count / total_scenarios if total_scenarios else 0.0,
            }
        )
    reason_df = pd.DataFrame(reason_rows)
    if not reason_df.empty:
        reason_df = reason_df.sort_values(["Type", "Count", "Driver"], ascending=[True, False, True]).reset_index(drop=True)

    breakpoints_df = critical_points["breakpoints_df"].copy()
    early_breaks = breakpoints_df[
        breakpoints_df["First non-resilient severity"].isin(["Light", "Base"])
    ].copy()

    if critical_count == total_scenarios and total_scenarios > 0:
        headline = (
            f"All {total_scenarios} scenarios are rated CRITICAL because every scenario crosses at least one hard distress trigger."
        )
    elif critical_count == 0 and watch_count > 0:
        headline = (
            f"No scenario is financially CRITICAL, but {watch_count} of {total_scenarios} scenarios fall below company benchmarks or turn cash-flow negative."
        )
    else:
        headline = (
            f"{critical_count} of {total_scenarios} scenarios are rated CRITICAL because they cross at least one hard distress trigger."
        )

    why_lines = []
    if base_distress_issues:
        why_lines.append("The base case already shows real financial stress before any scenario is applied:")
        why_lines.extend(f"- {line}" for line in base_distress_issues[:4])
    elif base_benchmark_gaps:
        why_lines.append("The base case is below the company's own benchmark on some metrics, but it is still above hard distress floors:")
        why_lines.extend(f"- {line}" for line in base_benchmark_gaps[:4])
    else:
        why_lines.append("The base case clears both the company benchmark and the hard distress floors; the break happens only after stress is applied.")

    critical_reason_df = reason_df[reason_df["Type"] == "Critical trigger"] if not reason_df.empty else pd.DataFrame()
    benchmark_reason_df = reason_df[reason_df["Type"] == "Benchmark pressure"] if not reason_df.empty else pd.DataFrame()
    if not critical_reason_df.empty:
        top_triggers = ", ".join(
            f"{row['Driver']} ({int(row['Count'])}/{total_scenarios})"
            for _, row in critical_reason_df.head(3).iterrows()
        )
        why_lines.append(f"When scenarios become CRITICAL, the main financial triggers are: {top_triggers}.")
    if not benchmark_reason_df.empty:
        top_benchmark_pressures = ", ".join(
            f"{row['Driver']} ({int(row['Count'])}/{total_scenarios})"
            for _, row in benchmark_reason_df.head(3).iterrows()
        )
        why_lines.append(f"Many scenarios also fall below the company's own benchmark levels: {top_benchmark_pressures}.")

    why_lines.append(
        "Below benchmark is not the same as distress. A company benchmark miss means the business is weaker than its own normal profile; CRITICAL means it crosses a hard danger line such as low coverage, excessive leverage, negative equity, negative cash, or deeply negative free cash flow."
    )

    look_at_lines = []
    if not critical_points["failure_df"].empty:
        for _, row in critical_points["failure_df"].head(3).iterrows():
            look_at_lines.append(
                f"- {row['Failure point']}: turns non-OK in {int(row['Fail count'])} of {total_scenarios} scenarios and fully critical in {int(row['Critical count'])}."
            )
    if not early_breaks.empty:
        first_rows = early_breaks.head(3)
        for _, row in first_rows.iterrows():
            look_at_lines.append(
                f"- {row['Sequence']}: first becomes non-resilient at {row['First non-resilient severity']}."
            )
    if not look_at_lines:
        look_at_lines.append("- No early breakpoints were identified under the current setup.")

    if base_distress_issues:
        decision_label = "Decide now"
        decision_text = (
            "The company already breaches a hard distress limit in the base case. That means this is not only a stress-case issue; the current financial profile itself needs immediate judgment."
        )
    elif light_critical > 0:
        decision_label = "Decide before committing capital"
        decision_text = (
            f"Light scenarios already turn critical ({light_critical} Light cases). "
            "This means small shocks are enough to push the company through a hard distress trigger. That is usually the point where the decision should move from interest to caution."
        )
    elif base_critical > 0:
        decision_label = "Do a deeper review before deciding"
        decision_text = (
            f"Base scenarios turn critical ({base_critical} Base cases), but Light scenarios do not. "
            "The company is not immediately fragile, but a normal stress case is enough to create genuine financial danger. The decision should depend on how likely you think that environment is."
        )
    elif severe_critical > 0:
        decision_label = "Tail-risk decision"
        decision_text = (
            f"Only Severe scenarios turn critical ({severe_critical} Severe cases). "
            "The decision is mainly about how much tail risk you are willing to accept."
        )
    elif watch_count > 0:
        decision_label = "Benchmark pressure, not distress"
        decision_text = (
            f"{watch_count} scenarios fall below company benchmarks, but they do not cross hard distress floors. "
            "This is a signal to review margin of safety, not an automatic rejection."
        )
    else:
        decision_label = "No immediate stress trigger"
        decision_text = (
            "The company does not breach hard distress rules or company benchmarks in the current scenario set. You can focus more on valuation, competitive position, and non-model risks."
        )

    return {
        "headline": headline,
        "why_text": "\n".join(why_lines),
        "look_at_text": "\n".join(look_at_lines),
        "decision_label": decision_label,
        "decision_text": decision_text,
        "base_status": base_status,
        "reason_df": reason_df,
    }


@st.cache_data(show_spinner=False)
def load_model():
    return load_workbook_model()


def load_company(symbol: str):
    return build_historical_dataset(symbol.upper().strip())


def _overview_score(overview: CompanyOverview | None) -> int:
    if overview is None:
        return -1
    score = 0
    if overview.current_price is not None:
        score += 1
    if overview.market_cap_m is not None:
        score += 1
    if bool((overview.summary or "").strip()):
        score += 1
    if overview.long_name and overview.long_name != overview.ticker:
        score += 1
    if overview.sector:
        score += 1
    if overview.industry:
        score += 1
    return score


def _merge_overviews(primary: CompanyOverview | None, secondary: CompanyOverview | None, ticker: str) -> CompanyOverview | None:
    if primary is None:
        return secondary
    if secondary is None:
        return primary

    def choose_text(primary_value: str, secondary_value: str, *, fallback_symbol: bool = False) -> str:
        cleaned_primary = (primary_value or "").strip()
        cleaned_secondary = (secondary_value or "").strip()
        if fallback_symbol and cleaned_primary == ticker and cleaned_secondary and cleaned_secondary != ticker:
            return cleaned_secondary
        return cleaned_primary or cleaned_secondary or (ticker if fallback_symbol else "")

    return CompanyOverview(
        ticker=ticker,
        short_name=choose_text(primary.short_name, secondary.short_name, fallback_symbol=True),
        long_name=choose_text(primary.long_name, secondary.long_name, fallback_symbol=True),
        sector=choose_text(primary.sector, secondary.sector),
        industry=choose_text(primary.industry, secondary.industry),
        exchange=choose_text(primary.exchange, secondary.exchange),
        currency=choose_text(primary.currency, secondary.currency),
        website=choose_text(primary.website, secondary.website),
        market_cap_m=primary.market_cap_m if primary.market_cap_m is not None else secondary.market_cap_m,
        current_price=primary.current_price if primary.current_price is not None else secondary.current_price,
        summary=choose_text(primary.summary, secondary.summary),
    )


def _repair_dataset_overview(dataset, active_ticker: str):
    session_overviews = st.session_state.setdefault("last_good_overviews", {})
    best_overview = _merge_overviews(dataset.overview, session_overviews.get(active_ticker), active_ticker)

    if _overview_score(best_overview) < 3:
        try:
            fresh_overview = fetch_company_overview(active_ticker)
        except Exception:
            fresh_overview = None
        best_overview = _merge_overviews(fresh_overview, best_overview, active_ticker)

    if best_overview is not None:
        previous_best = session_overviews.get(active_ticker)
        if _overview_score(best_overview) >= _overview_score(previous_best):
            session_overviews[active_ticker] = best_overview
        dataset.overview = session_overviews.get(active_ticker, best_overview)

    return dataset


def _first_quartile_series(series: pd.Series, lower: float | None = None, upper: float | None = None) -> float | None:
    clean = pd.to_numeric(series, errors="coerce").dropna()
    if clean.empty:
        return None
    if lower is not None or upper is not None:
        clean = clean.clip(lower=lower, upper=upper)
    return float(clean.quantile(0.25))


def _historical_current_ratio_series(financials: pd.DataFrame) -> pd.Series:
    current_assets = (
        pd.to_numeric(financials.loc["Cash & Equivalents"], errors="coerce").fillna(0.0)
        + pd.to_numeric(financials.loc["Accounts Receivable"], errors="coerce").fillna(0.0)
        + pd.to_numeric(financials.loc["Inventory"], errors="coerce").fillna(0.0)
        + pd.to_numeric(financials.loc["Other Current Assets"], errors="coerce").fillna(0.0)
    )
    current_liabilities = (
        pd.to_numeric(financials.loc["Accounts Payable"], errors="coerce").fillna(0.0)
        + pd.to_numeric(financials.loc["Other Current Liabilities"], errors="coerce").fillna(0.0)
        + pd.to_numeric(financials.loc["Short-term Debt"], errors="coerce").fillna(0.0)
    )
    return current_assets / current_liabilities.replace(0.0, pd.NA)


def build_threshold_calibration(dataset, defaults: ThresholdSettings) -> ThresholdCalibration:
    financials = dataset.financials
    notes: list[str] = []
    detail_rows: list[dict[str, str]] = []

    sga_ratio_series = pd.to_numeric(financials.loc["SG&A"], errors="coerce") / pd.to_numeric(
        financials.loc["Revenue"], errors="coerce"
    ).replace(0.0, pd.NA)
    q1_sga_share = _first_quartile_series(sga_ratio_series, lower=SGA_SHARE_LOWER_BOUND, upper=SGA_SHARE_UPPER_BOUND)
    sga_share = defaults.sga_variable_cost_share if q1_sga_share is None else q1_sga_share
    sga_share = min(max(float(sga_share), SGA_SHARE_LOWER_BOUND), SGA_SHARE_UPPER_BOUND)
    detail_rows.append(
        {
            "Setting": "SG&A variable cost share",
            "Historical 25th percentile": format_pct(q1_sga_share) if q1_sga_share is not None else "n/a",
            "Hard minimum / ceiling": f"{SGA_SHARE_LOWER_BOUND:.0%} to {SGA_SHARE_UPPER_BOUND:.0%} clamp",
            "Used in model": format_pct(sga_share),
            "Comment": "Uses the first quartile of SG&A / revenue as a company-specific proxy.",
        }
    )

    q1_cash_buffer = _first_quartile_series(pd.to_numeric(financials.loc["Cash & Equivalents"], errors="coerce"), lower=0.0)
    min_cash_buffer = HARD_MINIMUM_CASH_BUFFER if q1_cash_buffer is None else max(q1_cash_buffer, HARD_MINIMUM_CASH_BUFFER)
    cash_comment = "Uses the first quartile of the company's cash balance history."
    if q1_cash_buffer is None:
        cash_comment = "Cash history unavailable, so the hard minimum is used."
        notes.append("Could not calculate a first-quartile cash buffer from history; the model uses the hard minimum of 0.0.")
    elif q1_cash_buffer < HARD_MINIMUM_CASH_BUFFER:
        cash_comment = "The first-quartile cash balance is below the hard minimum, so the hard minimum is used."
        notes.append(
            f"The company does not satisfy the minimum cash buffer on a first-quartile basis; the model uses {format_m(HARD_MINIMUM_CASH_BUFFER)}."
        )
    detail_rows.append(
        {
            "Setting": "Minimum cash buffer",
            "Historical 25th percentile": format_m(q1_cash_buffer) if q1_cash_buffer is not None else "n/a",
            "Hard minimum / ceiling": format_m(HARD_MINIMUM_CASH_BUFFER),
            "Used in model": format_m(min_cash_buffer),
            "Comment": cash_comment,
        }
    )

    q1_interest_coverage = _first_quartile_series(
        pd.to_numeric(financials.loc["Interest Coverage"], errors="coerce"),
        lower=0.0,
        upper=25.0,
    )
    min_interest_coverage = (
        HARD_MINIMUM_INTEREST_COVERAGE
        if q1_interest_coverage is None
        else max(q1_interest_coverage, HARD_MINIMUM_INTEREST_COVERAGE)
    )
    interest_comment = "Uses the first quartile of historical EBIT / interest coverage."
    if q1_interest_coverage is None:
        interest_comment = "Coverage history unavailable, so the hard minimum is used."
        notes.append(
            f"Could not calculate a first-quartile EBIT / interest threshold; the model uses the hard minimum of {HARD_MINIMUM_INTEREST_COVERAGE:.2f}x."
        )
    elif q1_interest_coverage < HARD_MINIMUM_INTEREST_COVERAGE:
        interest_comment = "The first-quartile coverage is below the hard minimum, so the hard minimum is used."
        notes.append(
            f"The company does not satisfy the minimum EBIT / interest threshold on a first-quartile basis: {q1_interest_coverage:.2f}x vs hard minimum {HARD_MINIMUM_INTEREST_COVERAGE:.2f}x."
        )
    detail_rows.append(
        {
            "Setting": "Minimum EBIT / interest",
            "Historical 25th percentile": format_ratio(q1_interest_coverage) if q1_interest_coverage is not None else "n/a",
            "Hard minimum / ceiling": format_ratio(HARD_MINIMUM_INTEREST_COVERAGE),
            "Used in model": format_ratio(min_interest_coverage),
            "Comment": interest_comment,
        }
    )

    q1_net_leverage = _first_quartile_series(
        pd.to_numeric(financials.loc["Net Leverage"], errors="coerce"),
        lower=-5.0,
        upper=15.0,
    )
    max_net_leverage = HARD_MAXIMUM_NET_LEVERAGE if q1_net_leverage is None else min(q1_net_leverage, HARD_MAXIMUM_NET_LEVERAGE)
    leverage_comment = "Uses the first quartile of historical net debt / EBITDA."
    if q1_net_leverage is None:
        leverage_comment = "Leverage history unavailable, so the hard ceiling is used."
        notes.append(
            f"Could not calculate a first-quartile net debt / EBITDA threshold; the model uses the hard ceiling of {HARD_MAXIMUM_NET_LEVERAGE:.2f}x."
        )
    elif q1_net_leverage > HARD_MAXIMUM_NET_LEVERAGE:
        leverage_comment = "The first-quartile leverage is above the hard ceiling, so the hard ceiling is used."
        notes.append(
            f"The company does not satisfy the maximum net debt / EBITDA threshold on a first-quartile basis: {q1_net_leverage:.2f}x vs hard ceiling {HARD_MAXIMUM_NET_LEVERAGE:.2f}x."
        )
    detail_rows.append(
        {
            "Setting": "Maximum net debt / EBITDA",
            "Historical 25th percentile": format_ratio(q1_net_leverage) if q1_net_leverage is not None else "n/a",
            "Hard minimum / ceiling": format_ratio(HARD_MAXIMUM_NET_LEVERAGE),
            "Used in model": format_ratio(max_net_leverage),
            "Comment": leverage_comment,
        }
    )

    q1_current_ratio = _first_quartile_series(_historical_current_ratio_series(financials), lower=0.0, upper=10.0)
    min_current_ratio = HARD_MINIMUM_CURRENT_RATIO if q1_current_ratio is None else max(q1_current_ratio, HARD_MINIMUM_CURRENT_RATIO)
    current_ratio_comment = "Uses the first quartile of the historical current ratio."
    if q1_current_ratio is None:
        current_ratio_comment = "Current-ratio history unavailable, so the hard minimum is used."
        notes.append(
            f"Could not calculate a first-quartile current ratio threshold; the model uses the hard minimum of {HARD_MINIMUM_CURRENT_RATIO:.2f}x."
        )
    elif q1_current_ratio < HARD_MINIMUM_CURRENT_RATIO:
        current_ratio_comment = "The first-quartile current ratio is below the hard minimum, so the hard minimum is used."
        notes.append(
            f"The company does not satisfy the minimum current ratio threshold on a first-quartile basis: {q1_current_ratio:.2f}x vs hard minimum {HARD_MINIMUM_CURRENT_RATIO:.2f}x."
        )
    detail_rows.append(
        {
            "Setting": "Minimum current ratio",
            "Historical 25th percentile": format_ratio(q1_current_ratio) if q1_current_ratio is not None else "n/a",
            "Hard minimum / ceiling": format_ratio(HARD_MINIMUM_CURRENT_RATIO),
            "Used in model": format_ratio(min_current_ratio),
            "Comment": current_ratio_comment,
        }
    )

    settings = ThresholdSettings(
        sga_variable_cost_share=sga_share,
        minimum_cash_buffer=min_cash_buffer,
        minimum_interest_coverage=min_interest_coverage,
        maximum_net_leverage=max_net_leverage,
        minimum_current_ratio=min_current_ratio,
    )
    details = pd.DataFrame(detail_rows)
    return ThresholdCalibration(settings=settings, details=details, notes=notes)


def _sync_threshold_widget_state(active_ticker: str, calibration: ThresholdCalibration):
    if st.session_state.get("thresholds_initialized_for") == active_ticker:
        return
    st.session_state[THRESHOLD_WIDGET_KEYS["sga_variable_cost_share"]] = float(calibration.settings.sga_variable_cost_share)
    st.session_state[THRESHOLD_WIDGET_KEYS["minimum_cash_buffer"]] = float(calibration.settings.minimum_cash_buffer)
    st.session_state[THRESHOLD_WIDGET_KEYS["minimum_interest_coverage"]] = float(calibration.settings.minimum_interest_coverage)
    st.session_state[THRESHOLD_WIDGET_KEYS["maximum_net_leverage"]] = float(calibration.settings.maximum_net_leverage)
    st.session_state[THRESHOLD_WIDGET_KEYS["minimum_current_ratio"]] = float(calibration.settings.minimum_current_ratio)
    st.session_state["thresholds_initialized_for"] = active_ticker


def thresholds_from_sidebar(calibration: ThresholdCalibration, active_ticker: str) -> ThresholdSettings:
    st.sidebar.subheader("Thresholds")
    st.sidebar.caption(
        "These are company benchmark thresholds, auto-calibrated from the historical first quartile (25th percentile) and then checked against hard minimum floors / ceilings. "
        "Falling below them does not automatically mean financial distress. The model also checks hard distress floors separately when it assigns CRITICAL ratings."
    )

    _sync_threshold_widget_state(active_ticker, calibration)

    if calibration.notes:
        st.sidebar.markdown("**Threshold comments**")
        for note in calibration.notes:
            st.sidebar.caption(note)

    with st.sidebar.expander("How the thresholds were set", expanded=False):
        st.dataframe(calibration.details, use_container_width=True, hide_index=True)

    if st.sidebar.button("Reset to company-calibrated thresholds", use_container_width=True):
        st.session_state.pop("thresholds_initialized_for", None)
        _sync_threshold_widget_state(active_ticker, calibration)
        st.rerun()

    sga_variable_cost_share = st.sidebar.number_input(
        "SG&A variable cost share",
        min_value=0.0,
        max_value=1.0,
        step=0.05,
        format="%.2f",
        key=THRESHOLD_WIDGET_KEYS["sga_variable_cost_share"],
        help="Auto-set from the company's first-quartile SG&A / revenue ratio as a proxy, then clamped to a reasonable range.",
    )
    minimum_cash_buffer = st.sidebar.number_input(
        "Minimum cash buffer",
        step=50.0,
        format="%.1f",
        help="In the same currency units used in the app output, which is millions.",
        key=THRESHOLD_WIDGET_KEYS["minimum_cash_buffer"],
    )
    minimum_interest_coverage = st.sidebar.number_input(
        "Minimum EBIT / interest",
        min_value=0.0,
        step=0.25,
        format="%.2f",
        key=THRESHOLD_WIDGET_KEYS["minimum_interest_coverage"],
    )
    maximum_net_leverage = st.sidebar.number_input(
        "Maximum net debt / EBITDA",
        min_value=0.0,
        step=0.25,
        format="%.2f",
        key=THRESHOLD_WIDGET_KEYS["maximum_net_leverage"],
    )
    minimum_current_ratio = st.sidebar.number_input(
        "Minimum current ratio",
        min_value=0.0,
        step=0.10,
        format="%.2f",
        key=THRESHOLD_WIDGET_KEYS["minimum_current_ratio"],
    )
    return ThresholdSettings(
        sga_variable_cost_share=sga_variable_cost_share,
        minimum_cash_buffer=minimum_cash_buffer,
        minimum_interest_coverage=minimum_interest_coverage,
        maximum_net_leverage=maximum_net_leverage,
        minimum_current_ratio=minimum_current_ratio,
    )


def make_download_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def render_overview(dataset, dashboard_matrix: pd.DataFrame | None, stress_error: StressModelDataError | Exception | None = None):
    overview = dataset.overview
    if dashboard_matrix is not None:
        critical_count = int((dashboard_matrix["Rating"] == "CRITICAL").sum())
        watch_count = int((dashboard_matrix["Rating"] == "WATCH").sum())
        resilient_count = int((dashboard_matrix["Rating"] == "RESILIENT").sum())
        worst_rating = "CRITICAL" if critical_count else "WATCH" if watch_count else "RESILIENT"
    else:
        critical_count = None
        watch_count = None
        resilient_count = None
        worst_rating = "n/a"

    st.title("What If Stress Test Web App")
    st.caption(
        f"Workbook source: `{Path('WhatIf_StressTest_v4_Fixed.xlsx').name}`  |  "
        f"Live statement data: Yahoo Finance via `yahooquery`  |  "
        f"Units: {overview.currency or 'currency not reported'} millions"
    )

    name_col, meta_col = st.columns([2.4, 1.6])
    with name_col:
        st.subheader(f"{overview.long_name} ({overview.ticker})")
        meta_parts = [part for part in [overview.sector, overview.industry, overview.exchange] if part]
        if meta_parts:
            st.write(" | ".join(meta_parts))
        if overview.summary:
            st.write(overview.summary[:900] + ("..." if len(overview.summary) > 900 else ""))
    with meta_col:
        st.metric("Current price", f"{overview.current_price:,.2f}" if overview.current_price is not None else "n/a")
        st.metric("Market cap", f"{overview.market_cap_m:,.0f}m" if overview.market_cap_m is not None else "n/a")
        st.metric("Data quality", f"{dataset.data_quality_score:.1f}%")
        st.metric("Worst scenario rating", worst_rating)

    if dataset.sector_warning:
        st.warning(dataset.sector_warning)
    if dataset.blockers and stress_error is None:
        st.warning(
            "The latest annual statements are missing fields needed for a reliable stress model: "
            + ", ".join(dataset.blockers)
        )
    if stress_error is not None:
        if isinstance(stress_error, StressModelDataError):
            st.warning(
                "Stress scenarios are unavailable for this ticker because Yahoo Finance did not report enough latest annual fields: "
                + ", ".join(stress_error.missing_fields)
                + ". The mapped historical statements are still available below."
            )
        else:
            st.warning(
                "Stress scenarios could not be calculated for this ticker due to an unexpected modeling error. "
                "The mapped historical statements are still available below."
            )
    if dataset.warnings:
        with st.expander("Data mapping warnings", expanded=False):
            for warning in dataset.warnings:
                st.write(f"- {warning}")

    card1, card2, card3, card4 = st.columns(4)
    card1.metric("Critical scenarios", str(critical_count) if critical_count is not None else "n/a")
    card2.metric("Watch scenarios", str(watch_count) if watch_count is not None else "n/a")
    card3.metric("Resilient scenarios", str(resilient_count) if resilient_count is not None else "n/a")
    card4.metric("Latest fiscal year", dataset.latest_year)


def render_result_explanation(dataset, dashboard_matrix: pd.DataFrame, thresholds: ThresholdSettings):
    explanation = build_result_explanation(dataset, dashboard_matrix, thresholds)

    st.markdown("**What This Result Means**")
    st.info(explanation["headline"])

    decision_label_lower = explanation["decision_label"].lower()
    if explanation["decision_label"] == "Decide now":
        decision_background = "#f8d8d3"
    elif any(token in decision_label_lower for token in ["decide", "review", "tail-risk", "benchmark", "pressure"]):
        decision_background = "#f8e8c8"
    else:
        decision_background = "#d9efe1"
    col1, col2, col3 = st.columns([1.2, 1.2, 1.0])
    with col1:
        render_message_card("Why This Happens", explanation["why_text"], "#eef3f8")
    with col2:
        render_message_card("What To Look At First", explanation["look_at_text"], "#f7f7f2")
    with col3:
        render_message_card(explanation["decision_label"], explanation["decision_text"], decision_background)

    st.markdown("**Base Case: Company Benchmark vs Distress Context**")
    base_display = explanation["base_status"][
        [
            "Metric",
            "Base",
            "Company Benchmark",
            "Distress Zone",
            "Benchmark Status",
            "Distress Status",
            "Interpretation",
        ]
    ].copy()
    st.dataframe(base_display, use_container_width=True, hide_index=True)

    if not explanation["reason_df"].empty:
        st.markdown("**What Pushes Scenarios Into Watch Or Critical**")
        reason_display = explanation["reason_df"].copy()
        reason_display["Share of scenarios"] = reason_display["Share of scenarios"].map(format_pct)
        st.dataframe(reason_display, use_container_width=True, hide_index=True)


def render_stress_unavailable(stress_error: StressModelDataError | Exception | None):
    if isinstance(stress_error, StressModelDataError):
        st.warning("Stress scenarios are unavailable for this ticker.")
        st.write("Missing latest-year fields required by the model:")
        st.write(", ".join(stress_error.missing_fields))
        st.caption("The Historicals tab still shows the Yahoo-mapped statements and sources for review.")
        return

    st.error("Stress scenarios could not be calculated because of an unexpected error.")
    if stress_error is not None:
        st.code(str(stress_error))
    st.caption("The Historicals tab still shows the Yahoo-mapped statements and sources for review.")


def render_full_dashboard(dataset, dashboard_matrix: pd.DataFrame, thresholds: ThresholdSettings):
    st.subheader("Full Dashboard")

    critical_count = int((dashboard_matrix["Rating"] == "CRITICAL").sum())
    watch_count = int((dashboard_matrix["Rating"] == "WATCH").sum())
    resilient_count = int((dashboard_matrix["Rating"] == "RESILIENT").sum())
    total_count = len(dashboard_matrix)
    latest = dataset.latest_values

    card1, card2, card3, card4 = st.columns([1, 1, 1, 1])
    with card1:
        render_color_card("CRITICAL", str(critical_count), RATING_BG["CRITICAL"])
    with card2:
        render_color_card("WATCH", str(watch_count), RATING_BG["WATCH"], text_color="#13202f")
    with card3:
        render_color_card("RESILIENT", str(resilient_count), RATING_BG["RESILIENT"])
    with card4:
        render_color_card("Total Scenarios", str(total_count), "#5f6b7a")

    render_result_explanation(dataset, dashboard_matrix, thresholds)

    st.markdown("**Base Case Reference**")
    base_df = pd.DataFrame(
        [
            {
                "Revenue ($mm)": format_m(latest["Revenue"]),
                "EBITDA ($mm)": format_m(latest["EBITDA"]),
                "EBITDA Margin": format_pct(_safe_ratio(latest["EBITDA"], latest["Revenue"])),
                "Net Income ($mm)": format_m(latest["Net Income"]),
                "FCF ($mm)": format_m(latest["FCF (CFO - Capex)"]),
                "Cash ($mm)": format_m(latest["Cash & Equivalents"]),
                "Net Lev.": format_ratio(latest["Net Leverage"]),
                "Int Cov.": format_ratio(latest["Interest Coverage"]),
            }
        ]
    )
    st.dataframe(base_df, use_container_width=True, hide_index=True)

    st.markdown("**Scenario / Severity Matrix**")
    display = dashboard_matrix[
        [
            "Sequence",
            "Severity",
            "Revenue Δ",
            "EBITDA Δ",
            "EBITDA Margin (Stressed)",
            "Net Income (Stressed)",
            "FCF (Stressed)",
            "Ending Cash",
            "Net Leverage",
            "Interest Coverage",
            "Rating",
        ]
    ].copy()
    display = display.rename(
        columns={
            "Sequence": "Scenario",
            "EBITDA Margin (Stressed)": "EBITDA Margin %",
            "Net Income (Stressed)": "Net Income",
            "FCF (Stressed)": "FCF",
            "Ending Cash": "End Cash",
            "Net Leverage": "Net Lev.",
            "Interest Coverage": "Int Cov.",
            "Rating": "Outcome",
        }
    )

    display["Revenue Δ"] = display["Revenue Δ"].map(format_signed_m)
    display["EBITDA Δ"] = display["EBITDA Δ"].map(format_signed_m)
    display["EBITDA Margin %"] = display["EBITDA Margin %"].map(format_pct)
    display["Net Income"] = display["Net Income"].map(format_m)
    display["FCF"] = display["FCF"].map(format_m)
    display["End Cash"] = display["End Cash"].map(format_m)
    display["Net Lev."] = display["Net Lev."].map(format_ratio)
    display["Int Cov."] = display["Int Cov."].map(format_ratio)

    def _style_dashboard_row(row: pd.Series):
        styles = [""] * len(row)
        row_index = {name: idx for idx, name in enumerate(row.index)}

        severity_color = SEVERITY_BG.get(row["Severity"])
        if severity_color and "Severity" in row_index:
            styles[row_index["Severity"]] = f"background-color: {severity_color}; font-weight: 700;"

        rating_color = RATING_BG.get(row["Outcome"])
        if rating_color and "Outcome" in row_index:
            text_color = "#13202f" if row["Outcome"] == "WATCH" else "#ffffff"
            styles[row_index["Outcome"]] = f"background-color: {rating_color}; color: {text_color}; font-weight: 700;"

        if "Scenario" in row_index:
            styles[row_index["Scenario"]] = "font-weight: 700;"

        return styles

    styled = display.style.apply(_style_dashboard_row, axis=1)
    styled = styled.hide(axis="index")
    st.dataframe(styled, use_container_width=True, height=900)


def _safe_ratio(numerator: float, denominator: float, fallback: float | None = None) -> float | None:
    if denominator in (0, 0.0) or pd.isna(denominator):
        return fallback
    return numerator / denominator


def render_key_critical_points(dataset, dashboard_matrix: pd.DataFrame, thresholds: ThresholdSettings, sequence_map: pd.DataFrame):
    st.subheader("Key Critical Points")

    critical_points = build_critical_points(dataset, dashboard_matrix, thresholds)
    failure_df = critical_points["failure_df"]
    breakpoints_df = critical_points["breakpoints_df"].copy()
    worst_cases = critical_points["worst_cases"].copy()
    insights = critical_points["insights"]
    heatmap_df = critical_points["heatmap_df"]

    breakpoints_df["Non-resilient Order"] = breakpoints_df["First non-resilient severity"].map(SEVERITY_ORDER).fillna(99)
    breakpoints_df["Critical Order"] = breakpoints_df["First critical severity"].map(SEVERITY_ORDER).fillna(99)
    breakpoints_df["Worst Rating Order"] = breakpoints_df["Worst rating"].map(RATING_ORDER).fillna(99)

    early_breaks = breakpoints_df[
        breakpoints_df["First non-resilient severity"].isin(["Light", "Base"])
    ].sort_values(["Non-resilient Order", "Critical Order", "Worst Rating Order", "Sequence"])
    early_break_count = len(early_breaks)

    critical_breaks = breakpoints_df[
        breakpoints_df["First critical severity"].isin(["Light", "Base"])
    ].sort_values(["Critical Order", "Worst Rating Order", "Sequence"])
    critical_break_count = len(critical_breaks)

    top_failure_point = failure_df.iloc[0]["Failure point"] if not failure_df.empty else "None"
    worst_case_label = "None"
    if not worst_cases.empty:
        top_worst = worst_cases.iloc[0]
        worst_case_label = f"{top_worst['Sequence']} / {top_worst['Severity']}"

    card1, card2, card3, card4 = st.columns(4)
    card1.metric("Breaks by Light/Base", early_break_count)
    card2.metric("Critical by Light/Base", critical_break_count)
    card3.metric("Most common pressure point", top_failure_point)
    card4.metric("Worst live-data case", worst_case_label)

    if insights:
        st.markdown("**What matters most**")
        for insight in insights:
            st.write(f"- {insight}")

    render_failure_heatmap(heatmap_df)

    base_status, _, _ = build_base_threshold_status(dataset, thresholds)
    st.markdown("**Base Case: Benchmark vs Distress Context**")
    st.dataframe(
        base_status[
            [
                "Metric",
                "Base",
                "Company Benchmark",
                "Distress Zone",
                "Benchmark Status",
                "Distress Status",
                "Interpretation",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    detail = breakpoints_df.merge(
        sequence_map[["Sequence", "Cause -> Effect Chain"]],
        on="Sequence",
        how="left",
    )
    detail = detail.rename(columns={"Cause -> Effect Chain": "Chain"})

    st.markdown("**Small Changes That Break the Company**")
    if early_breaks.empty:
        st.info("No sequence becomes non-resilient at Light or Base severity under the current thresholds.")
    else:
        early_display = detail[detail["Sequence"].isin(early_breaks["Sequence"])][
            ["Sequence", "First non-resilient severity", "First critical severity", "Worst rating", "Chain"]
        ].sort_values(
            by=[
                "First non-resilient severity",
                "First critical severity",
                "Worst rating",
                "Sequence",
            ],
            key=lambda col: col.map({**SEVERITY_ORDER, **RATING_ORDER}).fillna(99) if col.name in {"First non-resilient severity", "First critical severity", "Worst rating"} else col,
        )
        st.dataframe(early_display, use_container_width=True, hide_index=True)

    st.markdown("**Most Common Pressure Points Across Scenarios**")
    if failure_df.empty:
        st.info("No dashboard pressure points were triggered under the current thresholds.")
    else:
        failure_display = failure_df.copy()
        failure_display["Share of scenarios"] = failure_display["Share of scenarios"].map(format_pct)
        failure_display = failure_display.rename(columns={"Fail count": "Non-OK count"})
        st.dataframe(failure_display, use_container_width=True, hide_index=True)

    st.markdown("**Worst Scenarios**")
    if worst_cases.empty:
        st.info("No stressed scenarios are available.")
    else:
        worst_display = worst_cases.copy()
        worst_display["Ending Cash"] = worst_display["Ending Cash"].map(format_m)
        worst_display["FCF (Stressed)"] = worst_display["FCF (Stressed)"].map(format_m)
        worst_display["Net Leverage"] = worst_display["Net Leverage"].map(format_ratio)
        worst_display["Interest Coverage"] = worst_display["Interest Coverage"].map(format_ratio)
        worst_display["Ending Equity"] = worst_display["Ending Equity"].map(format_m)
        st.dataframe(worst_display, use_container_width=True, hide_index=True)


def render_historical(dataset):
    st.subheader("Historical Financials")
    display_df = dataset.financials.copy().reset_index().rename(columns={"index": "Line Item"})
    ratio_rows = {
        "Gross Margin %",
        "EBITDA Margin %",
        "EBIT Margin %",
        "Net Margin %",
        "Revenue YoY %",
        "EBITDA YoY %",
        "Capex / Revenue %",
    }
    ratio_x_rows = {"Net Leverage", "Interest Coverage"}
    day_rows = {"DSO (days)", "DIO (days)", "DPO (days)"}

    for year in dataset.financials.columns:
        formatted = []
        for line_item, value in zip(display_df["Line Item"], display_df[year]):
            if pd.isna(value):
                formatted.append("n/a")
            elif line_item in ratio_rows:
                formatted.append(format_pct(value))
            elif line_item in ratio_x_rows:
                formatted.append(format_ratio(value))
            elif line_item in day_rows:
                formatted.append(format_days(value))
            else:
                formatted.append(format_m(value))
        display_df[year] = formatted
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    with st.expander("Source mapping by field and year", expanded=False):
        source_df = dataset.sources.copy().reset_index().rename(columns={"index": "Line Item"})
        st.dataframe(source_df, use_container_width=True, hide_index=True)


@st.cache_data(show_spinner=False, ttl=3600)
def _get_quarterly_chart_frame(symbol: str) -> tuple[pd.DataFrame, list[str]]:
    return build_quarterly_chart_frame(symbol)


def _build_dark_line_chart(
    frame: pd.DataFrame,
    series_specs: list[tuple[str, str, str, str]],
    title: str,
    yaxis_title: str,
    *,
    percent: bool = False,
) -> go.Figure:
    fig = go.Figure()

    for column, label, color, dash in series_specs:
        if column not in frame.columns:
            continue
        series = pd.to_numeric(frame[column], errors="coerce")
        if series.dropna().empty:
            continue

        hover_format = "%{y:.1%}" if percent else "%{y:,.1f}"
        fig.add_trace(
            go.Scatter(
                x=frame["asOfDate"],
                y=series,
                mode="lines+markers",
                name=label,
                line={"color": color, "width": 3, "dash": dash},
                marker={"size": 7},
                hovertemplate=f"%{{x|%b %Y}}<br>{label}: {hover_format}<extra></extra>",
            )
        )

        last_valid_index = series.last_valid_index()
        if last_valid_index is not None:
            value = series.loc[last_valid_index]
            label_text = f"{value:.0%}" if percent else f"{value:,.0f}"
            fig.add_trace(
                go.Scatter(
                    x=[frame.loc[last_valid_index, "asOfDate"]],
                    y=[value],
                    mode="text",
                    text=[label_text],
                    textposition="top center",
                    textfont={"color": "#d6d6d6", "size": 12},
                    showlegend=False,
                    hoverinfo="skip",
                )
            )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#222733",
        plot_bgcolor="#222733",
        title={"text": title, "x": 0.5, "xanchor": "center"},
        margin={"l": 20, "r": 20, "t": 60, "b": 50},
        height=430,
        hovermode="x unified",
        legend={"orientation": "h", "yanchor": "top", "y": -0.18, "xanchor": "left", "x": 0.0},
        yaxis={"title": yaxis_title, "tickformat": ".0%" if percent else ",.0f", "gridcolor": "#3b4151"},
        xaxis={"title": "", "tickformat": "%b\n%Y", "gridcolor": "#303744"},
    )
    return fig


def render_charts(dataset):
    st.subheader("Charts")
    st.caption(
        "These charts use Yahoo Finance quarterly (`3M`) statement data. "
        "Revenue, profit, assets, and liabilities are shown as quarterly values. "
        "ROIC, ROE, and ROA are shown as trailing-four-quarter ratios at each quarter-end."
    )

    symbol = getattr(getattr(dataset, "overview", None), "ticker", None)
    if not symbol:
        st.warning("The active ticker is unavailable, so the charts cannot be built.")
        return

    try:
        frame, quarterly_warnings = _get_quarterly_chart_frame(symbol)
    except Exception as exc:
        st.warning("Quarterly charts are temporarily unavailable for this ticker.")
        with st.expander("Quarterly chart error details", expanded=False):
            st.write(f"- {exc}")
        return

    if frame.empty:
        st.warning("Yahoo Finance did not return enough quarterly statement data to build the charts for this ticker.")
        return

    if quarterly_warnings:
        with st.expander("Quarterly data notes", expanded=False):
            for warning in quarterly_warnings:
                st.write(f"- {warning}")

    revenue_specs = [
        ("Revenue", "Revenue", "#2f7ed8", "solid"),
        ("Gross Profit", "Gross Profit", "#00d2d3", "solid"),
        ("EBIT", "EBIT", "#f5a25d", "solid"),
        ("Net Income", "Net Profit", "#2ecc71", "solid"),
    ]
    balance_specs = [
        ("Total Assets", "Total Assets", "#21ff3f", "solid"),
        ("Cash & Equivalents", "Cash", "#00d16a", "solid"),
        ("Current Assets", "Current assets", "#8fd35a", "solid"),
        ("Non-current Assets", "Non-current assets", "#567d1f", "solid"),
        ("Total Liabilities", "Total Liabilities", "#ff3b30", "dash"),
        ("Current Liabilities", "Current Liabilities", "#ff7b5a", "dash"),
        ("Non-current Liabilities", "Non-current Liabilities", "#ffb0a8", "dash"),
    ]
    returns_specs = [
        ("ROIC TTM %", "Return on Invested Capital", "#7f1d8d", "solid"),
        ("ROE TTM %", "Return on Equity", "#6d28d9", "solid"),
        ("ROA TTM %", "Return on Assets", "#c4a1ff", "solid"),
    ]
    margin_specs = [
        ("Gross Margin %", "Gross profit %", "#00c2b2", "solid"),
        ("EBIT Margin %", "EBIT Margin %", "#00a676", "solid"),
        ("Net Margin %", "Net income %", "#22c55e", "solid"),
    ]

    top_left, top_right = st.columns(2)
    with top_left:
        st.plotly_chart(
            _build_dark_line_chart(frame, revenue_specs, "#1 Revenue vs Profit", "Statement value ($mm)"),
            use_container_width=True,
        )
    with top_right:
        st.plotly_chart(
            _build_dark_line_chart(frame, balance_specs, "#6 Assets and Liabilities", "Balance sheet value ($mm)"),
            use_container_width=True,
        )

    bottom_left, bottom_right = st.columns(2)
    with bottom_left:
        st.plotly_chart(
            _build_dark_line_chart(frame, returns_specs, "#20 Profitability - ROIC - ROE - ROA", "TTM return", percent=True),
            use_container_width=True,
        )
    with bottom_right:
        st.plotly_chart(
            _build_dark_line_chart(frame, margin_specs, "#21 Profitability Margins", "Quarterly margin", percent=True),
            use_container_width=True,
        )

    with st.expander("Quarterly data used in charts", expanded=False):
        display_columns = [
            "Quarter Label",
            "Revenue",
            "Gross Profit",
            "EBIT",
            "Net Income",
            "Cash & Equivalents",
            "Current Assets",
            "Current Liabilities",
            "Total Assets",
            "Total Liabilities",
            "Gross Margin %",
            "EBIT Margin %",
            "Net Margin %",
            "ROIC TTM %",
            "ROE TTM %",
            "ROA TTM %",
        ]
        available_columns = [column for column in display_columns if column in frame.columns]
        display_frame = frame[available_columns].copy()
        for column in display_frame.columns:
            if column == "Quarter Label":
                continue
            if column.endswith("%"):
                display_frame[column] = display_frame[column].map(format_pct)
            else:
                display_frame[column] = display_frame[column].map(format_m)
        st.dataframe(display_frame, use_container_width=True, hide_index=True)


def render_ratio_scorecard(dataset):
    st.subheader("Financial Ratio Scorecard")
    st.caption(
        "Ratios and stars use only Yahoo Finance annual statement data mapped by the app. "
        "If Yahoo does not report the fields needed for a ratio, the app leaves it as `n/a`."
    )

    scorecard = build_ratio_scorecard(dataset)

    summary_cols = st.columns(len(CATEGORY_ORDER))
    for column, category in zip(summary_cols, CATEGORY_ORDER):
        with column:
            render_star_card(category, scorecard.summary_scores.get(category))

    st.markdown("**How the stars work**")
    st.markdown(
        """
- `max`: higher latest value versus the company's own Yahoo history gets more stars.
- `min`: lower latest value versus the company's own Yahoo history gets more stars.
- `avg`: values closer to the company's historical midpoint get more stars.
- Numeric marks such as `1-2`, `1`, `0.20`, or `>2.99` are scored against those explicit targets.
- Category stars are the average of the available ratio stars in that category.
- `Total score` is the average of the available category scores.
        """
    )

    if scorecard.notes:
        with st.expander("Data availability notes", expanded=False):
            for note in scorecard.notes:
                st.write(f"- {note}")

    with st.expander("Derived data used in the ratios", expanded=False):
        st.dataframe(scorecard.derived_table, use_container_width=True)

    with st.expander("Ratio formulas and mark logic", expanded=False):
        st.dataframe(scorecard.formula_table, use_container_width=True, hide_index=True)

    for category in ["Profitability", "Liquidity", "Credit risk", "Activity"]:
        score = scorecard.summary_scores.get(category)
        score_label = "n/a" if score is None or pd.isna(score) else f"{score:.1f}/5"
        st.markdown(f"**{category}**")
        st.caption(f"{stars_text(score)}  {score_label}")
        st.dataframe(scorecard.category_tables[category], use_container_width=True, hide_index=True)


def render_multiples_snapshot(dataset):
    st.subheader("Multiples")
    st.caption(
        "These are current market-value multiples and supporting quality metrics built from live Yahoo market data and the latest annual statements mapped in the app."
    )

    snapshot = build_multiples_snapshot(dataset)

    if snapshot.summary_cards:
        summary_cols = st.columns(len(snapshot.summary_cards))
        for column, card in zip(summary_cols, snapshot.summary_cards):
            with column:
                st.metric(card["title"], card["value"])

    if snapshot.notes:
        with st.expander("Data availability and interpretation notes", expanded=False):
            for note in snapshot.notes:
                st.write(f"- {note}")

    with st.expander("Inputs used in the multiples", expanded=False):
        st.dataframe(snapshot.inputs_table, use_container_width=True, hide_index=True)

    with st.expander("Formulas and methodology", expanded=False):
        st.dataframe(snapshot.formula_table, use_container_width=True, hide_index=True)

    for category in ["Market Value", "Market Multiples", "Quality & Capital Return"]:
        if category not in snapshot.category_tables:
            continue
        st.markdown(f"**{category}**")
        st.dataframe(snapshot.category_tables[category], use_container_width=True, hide_index=True)


def render_stock_scoring_model(dataset):
    st.subheader("Scoring Model")
    st.caption(
        "This scorecard adapts the stock-analysis spec into a live Yahoo Finance model for one ticker. "
        "It uses the company's own annual history and current Yahoo market data, without inserting fake industry benchmarks."
    )

    model = build_stock_scoring_model(dataset)

    summary_cols = st.columns(len(model.summary_cards))
    for column, card in zip(summary_cols, model.summary_cards):
        with column:
            st.metric(card["title"], card["value"])

    st.markdown("**How the model works**")
    st.markdown(
        """
- `Growth` uses the latest three Yahoo annual periods for revenue CAGR and earnings CAGR.
- `Profitability` scores operating margin, net margin, margin trend, and ROIC.
- `Financial Health` scores debt-to-equity, current ratio, quick ratio, and interest coverage.
- `Valuation` scores live market multiples and compares current P/E and P/S against the stock's own recent Yahoo annual valuation history.
- Each metric is scored on a `25 / 50 / 75 / 100` scale, category scores are averaged, and the final score uses `30% / 30% / 20% / 20%` weights.
- The recommendation is a rule-based output from the model, not investment advice.
        """
    )

    if model.notes:
        with st.expander("Data availability and model notes", expanded=False):
            for note in model.notes:
                st.write(f"- {note}")

    if not model.valuation_history_table.empty:
        with st.expander("Historical valuation context from Yahoo annual market-cap data", expanded=False):
            st.dataframe(model.valuation_history_table, use_container_width=True, hide_index=True)

    with st.expander("Inputs used in scoring", expanded=False):
        st.dataframe(model.inputs_table, use_container_width=True, hide_index=True)

    with st.expander("Formulas, thresholds, and weights", expanded=False):
        st.dataframe(model.formula_table, use_container_width=True, hide_index=True)

    for category in ["Growth", "Profitability", "Financial Health", "Valuation"]:
        score = model.category_scores.get(category)
        score_label = "n/a" if score is None or pd.isna(score) else f"{score:.1f}/100"
        st.markdown(f"**{category}**")
        st.caption(score_label)
        st.dataframe(model.category_tables[category], use_container_width=True, hide_index=True)


def render_selected_scenario(dataset, scenario_library: pd.DataFrame, thresholds: ThresholdSettings):
    st.subheader("Selected Scenario Dashboard")
    sequence_options = scenario_library["Sequence"].dropna().unique().tolist()
    default_sequence = "Demand collapse (classic recession)" if "Demand collapse (classic recession)" in sequence_options else sequence_options[0]
    if st.session_state.get("scenario_explorer_sequence") not in sequence_options:
        st.session_state["scenario_explorer_sequence"] = default_sequence
    if st.session_state.get("scenario_explorer_severity") not in {"Light", "Base", "Severe"}:
        st.session_state["scenario_explorer_severity"] = "Base"

    selector_col, severity_col, description_col = st.columns([1.6, 1.0, 2.4])
    with selector_col:
        selected_sequence = st.selectbox("Sequence", options=sequence_options, key="scenario_explorer_sequence")
    with severity_col:
        selected_severity = st.selectbox("Severity", options=["Light", "Base", "Severe"], key="scenario_explorer_severity")

    selected_row = get_selected_scenario(scenario_library, selected_sequence, selected_severity)
    selected_result = run_scenario(dataset.latest_values, selected_row, thresholds)

    with description_col:
        if st.session_state.get("scenario_explorer_source") == "failure_heatmap":
            focus_metric = st.session_state.get("scenario_explorer_metric_label", "Selected metric")
            st.caption(f"Preloaded from the failure heatmap. Current focus: `{focus_metric}`.")
        st.markdown("**Cause → effect chain**")
        st.write(selected_row["Cause -> Effect Chain"])
        st.markdown("**Rating**")
        st.write(f"{selected_result['Rating']} — {selected_result['Overall Assessment']}")
        if selected_result["Critical Reasons"] != "None":
            st.caption(f"Critical triggers: {selected_result['Critical Reasons']}")
        elif selected_result["Benchmark Reasons"] != "None":
            st.caption(f"Benchmark pressure: {selected_result['Benchmark Reasons']}")
        else:
            st.caption(selected_result["Rating Reasons"])

    metric_rows = [
        ("Revenue", dataset.latest_values["Revenue"], selected_result["Revenue (Stressed)"], None),
        ("EBITDA", dataset.latest_values["EBITDA"], selected_result["EBITDA (Stressed)"], None),
        ("Net income", dataset.latest_values["Net Income"], selected_result["Net Income (Stressed)"], None),
        ("FCF", dataset.latest_values["FCF (CFO - Capex)"], selected_result["FCF (Stressed)"], selected_result["Dashboard Status"]["FCF"]),
        ("Ending cash", dataset.latest_values["Cash & Equivalents"], selected_result["Ending Cash"], selected_result["Dashboard Status"]["Ending cash"]),
        ("Funding gap", 0.0, selected_result["Funding Gap"], selected_result["Dashboard Status"]["Funding gap"]),
        ("Net debt / EBITDA", dataset.latest_values["Net Leverage"], selected_result["Net Leverage"], selected_result["Dashboard Status"]["Net debt / EBITDA"]),
        ("EBIT / interest", dataset.latest_values["Interest Coverage"], selected_result["Interest Coverage"], selected_result["Dashboard Status"]["EBIT / interest"]),
        ("Current ratio", (dataset.latest_values["Cash & Equivalents"] + dataset.latest_values["Accounts Receivable"] + dataset.latest_values["Inventory"] + dataset.latest_values.get("Other Current Assets", 0.0)) / max(dataset.latest_values.get("Accounts Payable", 0.0) + dataset.latest_values.get("Other Current Liabilities", 0.0) + dataset.latest_values.get("Short-term Debt", 0.0), 0.01), selected_result["Current Ratio"], selected_result["Dashboard Status"]["Current ratio"]),
        ("Ending equity", dataset.latest_values["Equity"], selected_result["Ending Equity"], selected_result["Dashboard Status"]["Ending equity"]),
    ]

    dashboard_records = []
    for label, base_value, stress_value, status in metric_rows:
        delta = stress_value - base_value
        dashboard_records.append(
            {
                "Metric": label,
                "Base": base_value,
                "Stressed": stress_value,
                "Delta": delta,
                "Status": status or "INFO",
            }
        )
    dashboard_df = pd.DataFrame(dashboard_records)

    top_cards = st.columns(5)
    top_cards[0].metric("Scenario rating", selected_result["Rating"])
    top_cards[1].metric("Overall assessment", selected_result["Overall Assessment"])
    top_cards[2].metric("Revenue shock", format_pct(selected_result["Revenue Shock %"]))
    top_cards[3].metric("Interest shock", f"{selected_result['Interest Shock (bps)']:.0f} bps")
    top_cards[4].metric("One-off cash charge", format_pct(selected_result["One-off Cash Charge % Rev"]))

    ratio_metrics = {"Net debt / EBITDA", "EBIT / interest", "Current ratio"}

    def _fmt_metric(metric: str, value: float) -> str:
        if metric in ratio_metrics:
            return format_ratio(value)
        return format_m(value)

    formatted_dashboard = dashboard_df.copy()
    for column in ["Base", "Stressed", "Delta"]:
        formatted_dashboard[column] = [_fmt_metric(metric, value) for metric, value in zip(formatted_dashboard["Metric"], formatted_dashboard[column])]
    focus_metric_key = st.session_state.get("scenario_explorer_metric")
    if focus_metric_key and focus_metric_key in set(formatted_dashboard["Metric"]):
        styled_dashboard = formatted_dashboard.style.apply(
            lambda row: [
                "background-color: #fff6db; font-weight: 700;" if row["Metric"] == focus_metric_key else ""
                for _ in row
            ],
            axis=1,
        )
        st.dataframe(styled_dashboard, use_container_width=True, hide_index=True)
    else:
        st.dataframe(formatted_dashboard, use_container_width=True, hide_index=True)

    stress_detail = selected_result["Stress Detail"]
    chart_df = pd.DataFrame(
        {
            "Metric": ["Revenue", "EBITDA", "Net Income", "FCF", "Ending Cash"],
            "Base": [
                stress_detail["Base Revenue"],
                stress_detail["Base EBITDA"],
                stress_detail["Base Net Income"],
                stress_detail["Base CFO"] - dataset.latest_values["Capex"],
                stress_detail["Base Cash"],
            ],
            "Stressed": [
                stress_detail["Stressed Revenue"],
                stress_detail["Stressed EBITDA"],
                stress_detail["Stressed Net Income"],
                selected_result["FCF (Stressed)"],
                stress_detail["Ending Cash"],
            ],
        }
    )
    plot_df = chart_df.melt(id_vars="Metric", value_vars=["Base", "Stressed"], var_name="State", value_name="Value")
    fig = px.bar(
        plot_df,
        x="Metric",
        y="Value",
        color="State",
        barmode="group",
        color_discrete_map={"Base": "#46637f", "Stressed": "#c65c43"},
        title="Base vs Selected Stress Scenario",
    )
    fig.update_layout(height=420, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig, use_container_width=True)


def render_scenario_matrix(dataset, scenario_matrix: pd.DataFrame):
    st.subheader("Full Scenario Matrix")
    filter_col, sort_col = st.columns([1.2, 2.8])
    with filter_col:
        selected_ratings = st.multiselect("Ratings", ["CRITICAL", "WATCH", "RESILIENT"], default=["CRITICAL", "WATCH", "RESILIENT"])
        selected_severities = st.multiselect("Severities", ["Light", "Base", "Severe"], default=["Light", "Base", "Severe"])
    filtered = scenario_matrix[
        scenario_matrix["Rating"].isin(selected_ratings) & scenario_matrix["Severity"].isin(selected_severities)
    ].copy()

    if filtered.empty:
        st.info("No scenarios match the current filters.")
        return

    table_columns = [
        "Sequence",
        "Severity",
        "Rating",
        "Rating Reasons",
        "Revenue Δ",
        "EBITDA Δ",
        "Net Income (Stressed)",
        "FCF (Stressed)",
        "Ending Cash",
        "Net Leverage",
        "Interest Coverage",
        "Current Ratio",
        "Ending Equity",
    ]
    display = filtered[table_columns].copy()

    for column in ["Revenue Δ", "EBITDA Δ", "Net Income (Stressed)", "FCF (Stressed)", "Ending Cash", "Ending Equity"]:
        display[column] = display[column].map(format_m)
    for column in ["Net Leverage", "Interest Coverage", "Current Ratio"]:
        display[column] = display[column].map(format_ratio)

    st.dataframe(display, use_container_width=True, hide_index=True)

    scatter = px.scatter(
        filtered,
        x="Interest Coverage",
        y="Net Leverage",
        color="Rating",
        symbol="Severity",
        hover_name="Scenario Key",
        hover_data={
            "FCF (Stressed)": ":.1f",
            "Ending Cash": ":.1f",
            "Current Ratio": ":.2f",
        },
        color_discrete_map={"CRITICAL": "#a93439", "WATCH": "#d49232", "RESILIENT": "#2f7e54"},
        title="Scenario Risk Map: Coverage vs Net Leverage",
    )
    scatter.add_hline(y=DISTRESS_MAX_NET_LEVERAGE, line_dash="dash", line_color="#7f7f7f")
    scatter.add_vline(x=DISTRESS_MIN_INTEREST_COVERAGE, line_dash="dash", line_color="#7f7f7f")
    scatter.update_layout(height=520, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(scatter, use_container_width=True)

    download_col1, download_col2 = st.columns(2)
    with download_col1:
        st.download_button(
            "Download scenario matrix CSV",
            data=make_download_bytes(filtered.drop(columns=["Dashboard Status", "Stress Detail"], errors="ignore")),
            file_name=f"{dataset.overview.ticker.lower()}_scenario_matrix.csv",
            mime="text/csv",
        )
    with download_col2:
        historical_export = dataset.financials.reset_index().rename(columns={"index": "Line Item"})
        st.download_button(
            "Download mapped historicals CSV",
            data=make_download_bytes(historical_export),
            file_name=f"{dataset.overview.ticker.lower()}_historicals.csv",
            mime="text/csv",
        )


def render_sequence_library(sequence_map: pd.DataFrame):
    st.subheader("What-If Sequence Library")
    st.caption("This is loaded directly from the workbook so the app stays aligned with the Excel model.")
    st.dataframe(sequence_map, use_container_width=True, hide_index=True)


def render_faq():
    st.subheader("FAQ / How To Use")
    st.caption("A practical guide for interpreting the platform and using it in investment or credit work.")

    st.markdown("**Quick Start**")
    st.markdown(
        """
1. Enter a ticker and load the company from Yahoo Finance.
2. Review the `Historicals` tab to confirm the statements mapped correctly.
3. Open `Full Dashboard` to see how the company behaves across all scenario chains and severities.
4. Open `Critical Points` to find the combinations of small changes that break the company.
5. Use the thresholds in the sidebar to reflect your own credit or underwriting standards.
6. Use the result as a decision framework, not as a shortcut score.
        """
    )

    with st.expander("1. What did we build?", expanded=True):
        st.markdown(
            """
We built a company-level `What If...` stress-testing platform.

It takes a live ticker, pulls the latest annual financial statements from Yahoo Finance, maps them into the same logic as the Excel stress-test model, runs the full scenario library, and shows the results in a dashboard.

The model is built around cause-and-effect chains rather than a single statistic. Instead of saying only "this company scores 72," it asks questions like:

- What happens if revenue falls 20%?
- What happens if margin compresses while working capital absorbs cash?
- What happens if debt becomes more expensive or harder to refinance?
- At what point does the company move from resilient to watch to critical?
            """
        )

    with st.expander("2. Why do we use it?"):
        st.markdown(
            """
We use it because a company usually does not fail from one number moving in isolation.

Real distress comes from chains of events:

- demand falls
- fixed costs stay in place
- EBITDA drops
- leverage rises
- lenders tighten
- cash gets squeezed

Traditional ratio screens often miss that sequence. This platform helps you see how operating pressure, balance-sheet pressure, and liquidity pressure interact.
            """
        )

    with st.expander("3. How does it help make a decision?"):
        st.markdown(
            """
It improves decision-making by showing whether the investment case depends on fragile assumptions.

You can use it to answer questions such as:

- Is the company still acceptable if demand is only slightly weaker than expected?
- Does a mild margin miss create a real liquidity problem?
- Does the business have room to absorb a refinancing shock?
- Is the current valuation or rating ignoring a real break point?

This helps with position sizing, underwriting, watchlist monitoring, downside analysis, risk-reward framing, and deciding whether more diligence is needed before acting.
            """
        )

    with st.expander("4. What are the trigger points?"):
        st.markdown(
            """
Trigger points are the financial conditions that tell you the company is no longer operating safely under stress.

The main triggers in this platform are:

- ending cash falling below the chosen minimum cash buffer
- a funding gap opening up
- net debt / EBITDA rising above the selected leverage threshold
- EBIT / interest falling below the selected coverage threshold
- current ratio falling below the selected liquidity threshold
- ending equity turning negative
- free cash flow becoming deeply negative in a stress case

These are not abstract red flags. They are practical points where management flexibility, lender confidence, or equity value can deteriorate quickly.
            """
        )

    with st.expander("5. Which elements matter most?"):
        st.markdown(
            """
The most important elements are the ones that usually drive a company into stress:

- revenue sensitivity
- gross margin pressure
- fixed versus flexible operating costs
- working capital behavior such as receivables, inventory, and payables
- capex intensity
- debt load and refinancing profile
- interest burden
- cash starting position
- equity cushion

The model matters most when several of these move together. A company can survive one problem. It is the combination that usually matters.
            """
        )

    with st.expander("6. Should this be used for one company or for an industry?"):
        st.markdown(
            """
The primary use is for an individual company.

That is where the platform is strongest, because the outputs depend on the exact capital structure, margin profile, working capital setup, and cash position of that company.

It is also useful across an industry, but in a different way. For an industry view, use the same scenario logic across several companies and compare:

- which company breaks first
- which company keeps positive free cash flow longest
- which company has the weakest refinancing resilience
- which company has the smallest headroom before hitting trigger points

So the model is company-specific first, and peer-comparison second.
            """
        )

    with st.expander("7. Why is this not just another score?"):
        st.markdown(
            """
This is different from a generic score because it is scenario-driven and transparent.

A normal score compresses many things into one output and often hides the path that created it. This platform does the opposite:

- it shows the scenario chain
- it shows the driver inputs
- it shows how the statements change
- it shows which exact trigger breaks first

That means you can challenge the logic, adjust thresholds, and understand why a company becomes stressed. It is not a black box and it is not pretending that one number can summarize every risk.
            """
        )

    with st.expander("8. How do we get the rating?"):
        st.markdown(
            """
Each scenario receives a rating based on the stressed outputs.

`CRITICAL` means one or more major distress conditions are hit, such as:

- leverage above the selected limit
- interest coverage below the selected minimum
- negative ending equity
- very deep negative free cash flow

`RESILIENT` means the company still clears the key hurdles under that stress.

`WATCH` sits in between. It means the company may still function, but the margin of safety is thinner and the scenario deserves attention.

So the rating is not assigned from opinion. It is generated from the stressed financial statements and the trigger thresholds.
            """
        )

    with st.expander("9. What is the single most important thing to look for?"):
        st.markdown(
            """
The most important thing is not the most extreme scenario.

The most important thing is the first realistic combination of small changes that breaks the company.

That is where the edge usually is. If a business becomes stressed only under an extreme collapse, that may be acceptable. If it breaks under a mild demand slowdown plus working-capital pressure plus slightly higher rates, that is much more important for decision-making.

This is why the `Critical Points` tab matters so much. It shows where fragility begins, not just where catastrophe ends.
            """
        )

    with st.expander("10. Why is it important to use this platform?"):
        st.markdown(
            """
It is important because it turns raw financial statements into a forward-looking stress framework.

Most investors and analysts can read income statements, balance sheets, and cash flows. Fewer can quickly translate those statements into answers to questions like:

- How much stress can this company absorb?
- What breaks first?
- How fast can liquidity disappear?
- Is this downside already visible in the current market view?

This platform helps you move from observation to decision. It makes downside analysis structured, repeatable, explainable, and fast enough to use across many names.
            """
        )

    st.markdown("**Interpretation Notes**")
    st.markdown(
        """
- Use the platform as a decision aid, not a substitute for judgment.
- Always review the `Historicals` and `Source mapping` sections before trusting the output.
- The model is strongest for operating companies and less suitable for banks, insurers, and other financial institutions.
- The adjustable thresholds matter. Different investors, lenders, and credit teams will define stress differently.
        """
    )


def main():
    scenario_library, sequence_map, default_thresholds = load_model()

    st.sidebar.title("Ticker Loader")
    if "ticker" not in st.session_state:
        st.session_state["ticker"] = "GM"
    if "ticker_input" not in st.session_state:
        st.session_state["ticker_input"] = st.session_state["ticker"]

    ticker_input = st.sidebar.text_input(
        "Yahoo Finance ticker",
        key="ticker_input",
        help="Examples: GM, AAPL, KO, BMW.DE",
    )
    load_pressed = st.sidebar.button("Load ticker", type="primary", use_container_width=True)
    if load_pressed:
        st.session_state["ticker"] = ticker_input.upper().strip()
        st.session_state.pop("thresholds_initialized_for", None)
        st.rerun()

    active_ticker = st.session_state["ticker"]

    try:
        with st.spinner(f"Loading live Yahoo Finance data for {active_ticker}..."):
            dataset = load_company(active_ticker)
    except Exception as exc:
        st.title("What If Stress Test Web App")
        st.error(f"Could not load Yahoo Finance data for `{active_ticker}`.")
        st.code(str(exc))
        st.stop()

    dataset = _repair_dataset_overview(dataset, active_ticker)

    calibration = build_threshold_calibration(dataset, default_thresholds)
    thresholds = thresholds_from_sidebar(calibration, active_ticker)

    scenario_matrix = None
    dashboard_matrix = None
    stress_error = None
    try:
        dataset.latest_values = prepare_latest_for_stress(dataset.latest_values)
        scenario_matrix = run_all_scenarios(dataset.latest_values, thresholds, scenario_library)
        dashboard_matrix = get_dashboard_matrix(scenario_matrix, scenario_library)
    except StressModelDataError as exc:
        stress_error = exc
    except Exception as exc:
        stress_error = exc

    render_overview(dataset, dashboard_matrix, stress_error)

    tabs = st.tabs(
        [
            "Full Dashboard",
            "Critical Points",
            "Scenario Explorer",
            "Financial Ratios",
            "Multiples",
            "Scoring Model",
            "Charts",
            "Historicals",
            "Scenario Matrix",
            "Sequence Map",
            "FAQ / How To Use",
        ]
    )
    with tabs[0]:
        if dashboard_matrix is None:
            render_stress_unavailable(stress_error)
        else:
            render_full_dashboard(dataset, dashboard_matrix, thresholds)
    with tabs[1]:
        if dashboard_matrix is None:
            render_stress_unavailable(stress_error)
        else:
            render_key_critical_points(dataset, dashboard_matrix, thresholds, sequence_map)
    with tabs[2]:
        if dashboard_matrix is None:
            render_stress_unavailable(stress_error)
        else:
            render_selected_scenario(dataset, scenario_library, thresholds)
    with tabs[3]:
        render_ratio_scorecard(dataset)
    with tabs[4]:
        render_multiples_snapshot(dataset)
    with tabs[5]:
        render_stock_scoring_model(dataset)
    with tabs[6]:
        render_charts(dataset)
    with tabs[7]:
        render_historical(dataset)
    with tabs[8]:
        if scenario_matrix is None:
            render_stress_unavailable(stress_error)
        else:
            render_scenario_matrix(dataset, scenario_matrix)
    with tabs[9]:
        render_sequence_library(sequence_map)
    with tabs[10]:
        render_faq()


if __name__ == "__main__":
    main()
