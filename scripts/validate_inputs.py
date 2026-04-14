import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from petrocaf_pricing.io.config_loader import load_config
from petrocaf_pricing.io.csv_io import read_csv, write_csv
from petrocaf_pricing.services.master_data_service import load_master_data
from petrocaf_pricing.services.validation_service import validate_boq
from petrocaf_pricing.utils.path_utils import resolve_path

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    args = parser.parse_args()
    config, project_root = load_config(args.config)
    boq = read_csv(str(resolve_path(project_root, config["input"]["boq_csv"])))
    master_paths = {k: resolve_path(project_root, v) for k, v in config["master_data"].items()}
    master_data = load_master_data(master_paths)
    report = validate_boq(
        boq,
        list(master_data.discipline_rules["discipline"].astype(str).str.lower().unique()),
        bool(config["pricing_options"].get("allow_zero_rates", False)),
    )
    out = resolve_path(project_root, config["output"]["validation_report_csv"])
    out.parent.mkdir(parents=True, exist_ok=True)
    write_csv(report, str(out))
    print(f"Validation report written to: {out}")
