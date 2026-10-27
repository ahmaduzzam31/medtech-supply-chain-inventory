#!/usr/bin/env bash
# Rebuild everything: simulated data -> Excel model -> reconciliation -> notebook -> dashboard.
set -euo pipefail
cd "$(dirname "$0")"
python src/generate_data.py
python src/build_workbook.py
python src/reconcile.py
jupyter nbconvert --to notebook --execute --inplace notebooks/inventory_analysis.ipynb
python src/build_dashboard.py
