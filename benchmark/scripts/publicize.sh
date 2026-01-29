#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
bench_dir="$(dirname "$script_dir")"

python_bin="$bench_dir/.venv/bin/python"
if [[ -f "$python_bin" ]]; then
  "$python_bin" "$script_dir/publicize.py"
else
  python "$script_dir/publicize.py"
fi
