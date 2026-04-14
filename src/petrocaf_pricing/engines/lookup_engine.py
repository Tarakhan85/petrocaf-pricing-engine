import pandas as pd

def lookup_scalar(df: pd.DataFrame, filter_col: str, filter_value: str, value_col: str, default: float = 1.0) -> float:
    rows = df[df[filter_col].astype(str).str.lower() == str(filter_value).lower()]
    if rows.empty:
        return default
    return float(rows.iloc[0][value_col])

def lookup_productivity(df: pd.DataFrame, discipline: str, unit: str, default: float = 1.0) -> float:
    rows = df[
        (df["discipline"].astype(str).str.lower() == str(discipline).lower()) &
        (df["unit"].astype(str).str.lower() == str(unit).lower())
    ]
    if rows.empty:
        return default
    return float(rows.iloc[0]["standard_output_per_day"])
