import sqlite3
from pathlib import Path
import pandas as pd

def build_sqlite(master_folder: str, db_path: str) -> str:
    master = Path(master_folder)
    db = Path(db_path)
    db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db)
    try:
        for csv_file in master.glob("*.csv"):
            df = pd.read_csv(csv_file)
            df.to_sql(csv_file.stem, conn, if_exists="replace", index=False)
        conn.commit()
    finally:
        conn.close()
    return str(db)
