import pandas as pd

REQUIRED_COLUMNS = [
    "item_code","description","discipline","unit","quantity",
    "base_material_rate","base_labor_rate","base_equipment_rate",
    "environment_factor","complexity_factor","location_factor"
]

def validate_boq(boq: pd.DataFrame, allowed_disciplines: list[str], allow_zero_rates: bool = False) -> pd.DataFrame:
    issues = []
    for col in REQUIRED_COLUMNS:
        if col not in boq.columns:
            issues.append({"level":"ERROR","item_code":"","field":col,"message":f"Missing required column: {col}"})
    if issues:
        return pd.DataFrame(issues)

    allowed = [d.lower() for d in allowed_disciplines]
    for idx, row in boq.iterrows():
        item = str(row.get("item_code", f"ROW-{idx+1}"))
        if str(row["discipline"]).lower() not in allowed:
            issues.append({"level":"ERROR","item_code":item,"field":"discipline","message":f"Unsupported discipline: {row['discipline']}"})
        if float(row["quantity"]) <= 0:
            issues.append({"level":"ERROR","item_code":item,"field":"quantity","message":"Quantity must be > 0"})
        for fld in ["base_material_rate","base_labor_rate","base_equipment_rate"]:
            val = float(row[fld])
            if val < 0:
                issues.append({"level":"ERROR","item_code":item,"field":fld,"message":"Rate cannot be negative"})
            if (not allow_zero_rates) and val == 0:
                issues.append({"level":"WARNING","item_code":item,"field":fld,"message":"Zero rate detected"})
        for fld in ["environment_factor","complexity_factor","location_factor"]:
            if float(row[fld]) <= 0:
                issues.append({"level":"ERROR","item_code":item,"field":fld,"message":"Factor must be > 0"})
    return pd.DataFrame(issues) if issues else pd.DataFrame(columns=["level","item_code","field","message"])
