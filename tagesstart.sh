#!/bin/bash
cd "$HOME/Jobsuche" || exit 1
PY="$HOME/Jobsuche/.venv/bin/python"
echo "===== $(date '+%d.%m.%Y %H:%M') ====="
"$PY" mailcheck.py
"$PY" jobfinder.py
