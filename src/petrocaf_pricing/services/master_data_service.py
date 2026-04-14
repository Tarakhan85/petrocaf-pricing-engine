from dataclasses import dataclass
from pathlib import Path
import pandas as pd
from petrocaf_pricing.io.csv_io import read_csv

@dataclass
class MasterData:
    discipline_rules: pd.DataFrame
    coefficients: pd.DataFrame
    productivities: pd.DataFrame
    indirects: pd.DataFrame
    markups: pd.DataFrame
    scenario_factors: pd.DataFrame

def load_master_data(paths: dict[str, Path]) -> MasterData:
    return MasterData(
        discipline_rules=read_csv(str(paths["discipline_rules"])),
        coefficients=read_csv(str(paths["coefficients"])),
        productivities=read_csv(str(paths["productivities"])),
        indirects=read_csv(str(paths["indirects"])),
        markups=read_csv(str(paths["markups"])),
        scenario_factors=read_csv(str(paths["scenario_factors"])),
    )
