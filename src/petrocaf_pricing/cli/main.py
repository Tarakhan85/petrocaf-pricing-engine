import argparse
from petrocaf_pricing.core.workflow import execute

def main():
    parser = argparse.ArgumentParser(description="PETROCAF pricing CLI")
    parser.add_argument("--config", required=True, help="Path to config JSON")
    args = parser.parse_args()
    print(execute(args.config))
