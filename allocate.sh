#!/usr/bin/env bash
# Usage: ./allocate.sh path/to/team_formation.xlsx [--deterministic --seed 1 ...]
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
[ -x "$here/.venv/bin/python" ] || { echo "First run ./setup_mac_linux.sh once."; exit 1; }
exec "$here/.venv/bin/python" "$here/allocate_teams.py" "$@"
