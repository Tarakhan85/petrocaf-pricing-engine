# PETROCAF Pricing Engine Full V1

GitHub-ready deterministic pricing engine for EPC / construction tender pricing.

This repository now includes a complete Excel workbook generator for PETROCAF pricing and BOQ cost build-up.

## Fastest run: generate the Excel pricing workbook

```bash
pip install openpyxl
python tools/generate_petrocaf_pricing_workbook.py
```

Output:

```text
PETROCAF_Pricing_Master_Workbook.xlsx
```

## Existing package quick start

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python scripts/run_pricing.py --config config/settings.json
```

## Workbook sheets generated

1. `README_BASIS` - pricing basis, governance and estimator instructions.
2. `PROJECT_MASTER` - project metadata, default OH, profit, contingency and tax placeholders.
3. `BOQ_INPUT` - BOQ input table with activity mapping fields.
4. `ACTIVITY_LIBRARY` - activity dictionary for piping, hydrotest, tank repair, civil, E&I and coating.
5. `PRODUCTIVITY_LIBRARY` - productivity assumptions with source status and adjustment factors.
6. `CREW_LIBRARY` - crew templates, manpower quantities, daily rates and manhours.
7. `MATERIAL_LIBRARY` - material and consumable cost placeholders.
8. `EQUIPMENT_LIBRARY` - equipment spreads and daily cost assumptions.
9. `INDIRECTS_RISK` - indirects, HSE/QAQC, mobilization, contingency, overhead and profit.
10. `EGYPT_CALIBRATION` - labor, material, equipment, productivity, logistics and risk factors.
11. `COST_ENGINE` - auditable calculation sheet linking BOQ, activity, productivity, crew, equipment, material, indirects, risk, OH and profit.
12. `OUTPUT_SUMMARY` - executive pricing summary and charts.
13. `REVIEW_FLAGS` - pre-submission checks for missing data and zero-cost issues.
14. `SOURCE_REGISTER` - evidence trail for rates, productivity and assumptions.
15. `CHANGE_LOG` - revision and approval log.

## Core commercial rules

- No fabricated source-backed pricing data.
- Unverified rates remain tagged as `ESTIMATED` or `MISSING`.
- Replace placeholders with PETROCAF-approved norms, supplier quotations or project-specific assumptions before any tender submission.
- AI may assist classification, review and gap detection only. Final pricing logic remains deterministic, auditable and rule-based.

## Intended scopes

The generated workbook is structured for EPC / oil & gas / industrial construction pricing, including:

- Mechanical piping fabrication and erection.
- Hydrotest packages.
- Pipe supports.
- Tank repair placeholders.
- Civil concrete and earthwork placeholders.
- Electrical and instrumentation placeholders.
- Painting/coating placeholders.

## Important warning

This workbook is a pricing engine template, not a final commercial offer. Before live use, review client specifications, site conditions, inclusions/exclusions, taxes, logistics, procurement quotations, schedule, productivity basis and payment terms.
