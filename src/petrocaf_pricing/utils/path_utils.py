from pathlib import Path

def resolve_path(project_root: Path, path_str: str) -> Path:
    p = Path(path_str)
    return p if p.is_absolute() else (project_root / p)

def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
