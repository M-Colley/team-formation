#!/usr/bin/env bash
# One-time setup: creates a private Python environment (.venv) in this folder
# and installs the packages from requirements.txt. Nothing else on the machine is changed.
set -euo pipefail
cd "$(dirname "$0")"
command -v python3 >/dev/null || { echo "Python 3.10+ is needed: https://www.python.org/downloads/"; exit 1; }
[ -d .venv ] || python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip >/dev/null
.venv/bin/python -m pip install -r requirements.txt
echo "Setup complete. Run: ./allocate.sh path/to/team_formation.xlsx"
