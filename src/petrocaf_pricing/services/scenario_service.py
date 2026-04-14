import pandas as pd

def get_scenario_row(scenario_df: pd.DataFrame, scenario_name: str) -> pd.Series:
    rows = scenario_df[scenario_df["scenario_name"].astype(str).str.lower() == str(scenario_name).lower()]
    if rows.empty:
        rows = scenario_df[scenario_df["scenario_name"].astype(str).str.lower() == "base"]
    return rows.iloc[0]
