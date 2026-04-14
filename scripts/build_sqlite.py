import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
from petrocaf_pricing.data_sqlite_builder import build_sqlite

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--master-folder", default="data/master")
    parser.add_argument("--db-path", default="data/output/petrocaf_master.db")
    args = parser.parse_args()
    print(build_sqlite(str(ROOT / args.master_folder), str(ROOT / args.db_path)))
