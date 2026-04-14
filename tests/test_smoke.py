import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
from petrocaf_pricing.core.workflow import execute

def test_smoke():
    result = execute("config/settings.json")
    assert result["status"] == "success"
