import json
from pathlib import Path

def load_config(config_path: str) -> tuple[dict, Path]:
    path = Path(config_path).resolve()
    with path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    return cfg, path.parent.parent
