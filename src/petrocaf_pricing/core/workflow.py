from petrocaf_pricing.io.config_loader import load_config
from petrocaf_pricing.io.csv_io import read_csv, write_csv
from petrocaf_pricing.services.master_data_service import load_master_data
from petrocaf_pricing.services.validation_service import validate_boq
from petrocaf_pricing.engines.pricing_engine import run_pricing
from petrocaf_pricing.utils.path_utils import resolve_path, ensure_parent
from petrocaf_pricing.utils.logging_utils import get_logger

logger = get_logger()

def execute(config_path: str):
    config, project_root = load_config(config_path)
    boq_path = resolve_path(project_root, config["input"]["boq_csv"])
    output_priced = resolve_path(project_root, config["output"]["priced_boq_csv"])
    output_summary = resolve_path(project_root, config["output"]["pricing_summary_csv"])
    output_validation = resolve_path(project_root, config["output"]["validation_report_csv"])
    master_paths = {k: resolve_path(project_root, v) for k, v in config["master_data"].items()}
    options = config["pricing_options"]

    boq = read_csv(str(boq_path))
    master_data = load_master_data(master_paths)

    allowed_disciplines = list(master_data.discipline_rules["discipline"].astype(str).str.lower().unique())
    validation = validate_boq(boq, allowed_disciplines, bool(options.get("allow_zero_rates", False)))
    ensure_parent(output_validation)
    write_csv(validation, str(output_validation))

    if not validation.empty and (validation["level"].astype(str).str.upper() == "ERROR").any():
        return {
            "status": "failed",
            "priced_boq_path": None,
            "summary_path": None,
            "validation_path": str(output_validation),
            "message": "Validation failed. Fix ERROR rows in validation report."
        }

    priced_boq, summary = run_pricing(
        boq=boq,
        master_data=master_data,
        scenario_name=str(options.get("scenario_name", "base")),
        rounding=int(options.get("default_rounding", 2)),
    )

    ensure_parent(output_priced)
    ensure_parent(output_summary)
    write_csv(priced_boq, str(output_priced))
    write_csv(summary, str(output_summary))

    return {
        "status": "success",
        "priced_boq_path": str(output_priced),
        "summary_path": str(output_summary),
        "validation_path": str(output_validation),
        "message": "Pricing completed successfully."
    }
