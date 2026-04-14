import pandas as pd
from petrocaf_pricing.engines.lookup_engine import lookup_scalar, lookup_productivity
from petrocaf_pricing.services.scenario_service import get_scenario_row
from petrocaf_pricing.utils.math_utils import safe_round

def run_pricing(boq: pd.DataFrame, master_data, scenario_name: str = "base", rounding: int = 2):
    result = boq.copy()

    coef = master_data.coefficients
    rules = master_data.discipline_rules
    prods = master_data.productivities
    indirects = master_data.indirects
    markups = master_data.markups
    scenario_row = get_scenario_row(master_data.scenario_factors, scenario_name)

    result["material_factor"] = result["discipline"].map(lambda d: lookup_scalar(coef, "discipline", d, "material_factor", 1.0))
    result["labor_factor"] = result["discipline"].map(lambda d: lookup_scalar(coef, "discipline", d, "labor_factor", 1.0))
    result["equipment_factor"] = result["discipline"].map(lambda d: lookup_scalar(coef, "discipline", d, "equipment_factor", 1.0))
    result["waste_factor"] = result["discipline"].map(lambda d: lookup_scalar(rules, "discipline", d, "default_waste_factor", 1.0))
    result["crew_factor"] = result["discipline"].map(lambda d: lookup_scalar(rules, "discipline", d, "default_crew_factor", 1.0))
    result["productivity_factor"] = result["discipline"].map(lambda d: lookup_scalar(rules, "discipline", d, "default_productivity_factor", 1.0))
    result["standard_output_per_day"] = result.apply(lambda r: lookup_productivity(prods, r["discipline"], r["unit"], 1.0), axis=1)
    result["estimated_duration_days"] = result["quantity"] / (result["standard_output_per_day"] * result["productivity_factor"])
    result["difficulty_factor"] = result["environment_factor"] * result["complexity_factor"] * result["location_factor"]

    mm = float(scenario_row["material_multiplier"])
    lm = float(scenario_row["labor_multiplier"])
    em = float(scenario_row["equipment_multiplier"])
    im = float(scenario_row["indirect_multiplier"])
    mkm = float(scenario_row["markup_multiplier"])

    result["material_cost"] = result["quantity"] * result["base_material_rate"] * result["material_factor"] * result["waste_factor"] * mm
    result["labor_cost"] = result["quantity"] * result["base_labor_rate"] * result["labor_factor"] * result["crew_factor"] * lm
    result["equipment_cost"] = result["quantity"] * result["base_equipment_rate"] * result["equipment_factor"] * em
    result["direct_cost_raw"] = result["material_cost"] + result["labor_cost"] + result["equipment_cost"]
    result["direct_cost"] = result["direct_cost_raw"] * result["difficulty_factor"]

    indirect_base = float(indirects["factor"].sum()) * im
    markup_base = float(markups["factor"].sum()) * mkm

    result["indirect_cost"] = result["direct_cost"] * indirect_base
    result["subtotal_before_markup"] = result["direct_cost"] + result["indirect_cost"]
    result["markup_cost"] = result["subtotal_before_markup"] * markup_base
    result["final_cost"] = result["subtotal_before_markup"] + result["markup_cost"]
    result["unit_rate"] = result["final_cost"] / result["quantity"]
    result["scenario_name"] = scenario_name

    for col in ["standard_output_per_day","estimated_duration_days","difficulty_factor","material_cost","labor_cost","equipment_cost","direct_cost_raw","direct_cost","indirect_cost","subtotal_before_markup","markup_cost","final_cost","unit_rate"]:
        result[col] = result[col].map(lambda x: safe_round(x, rounding))

    summary = pd.DataFrame([{
        "scenario_name": scenario_name,
        "items_count": int(len(result)),
        "total_quantity": safe_round(result["quantity"].sum(), rounding),
        "total_direct_cost": safe_round(result["direct_cost"].sum(), rounding),
        "total_indirect_cost": safe_round(result["indirect_cost"].sum(), rounding),
        "total_markup_cost": safe_round(result["markup_cost"].sum(), rounding),
        "total_final_cost": safe_round(result["final_cost"].sum(), rounding),
        "average_unit_rate": safe_round(result["unit_rate"].mean(), rounding),
        "estimated_total_duration_days": safe_round(result["estimated_duration_days"].sum(), rounding),
    }])
    return result, summary
