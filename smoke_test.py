from __future__ import annotations

from financial_ratios import build_ratio_scorecard
from multiples import build_multiples_snapshot
from stress_backend import StressModelDataError, build_historical_dataset, load_workbook_model, run_all_scenarios


def main() -> None:
    scenario_library, _, defaults = load_workbook_model()
    for symbol in ["GM", "AAPL", "KO"]:
        dataset = build_historical_dataset(symbol)
        matrix = run_all_scenarios(dataset.latest_values, defaults, scenario_library)
        scorecard = build_ratio_scorecard(dataset)
        multiples = build_multiples_snapshot(dataset)
        counts = matrix["Rating"].value_counts().to_dict()
        print(symbol, "latest_year", dataset.latest_year, "quality", dataset.data_quality_score, "blockers", dataset.blockers)
        print(symbol, "ratings", counts)
        print(symbol, "ratio_total_score", scorecard.summary_scores.get("Total score"))
        if multiples.summary_cards:
            print(symbol, "pe", multiples.summary_cards[0]["value"])

    for symbol in ["P911.DE", "CHA"]:
        dataset = build_historical_dataset(symbol)
        try:
            run_all_scenarios(dataset.latest_values, defaults, scenario_library)
        except StressModelDataError as exc:
            print(symbol, "stress unavailable", exc.missing_fields)
        else:
            print(symbol, "stress available", "blockers", dataset.blockers)


if __name__ == "__main__":
    main()
