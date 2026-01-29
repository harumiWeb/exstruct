#!/usr/bin/env bash
set -euo pipefail

CASE="all"
METHOD="all"
MODEL="gpt-4o"
TEMPERATURE="0.0"
SKIP_ASK="false"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --case) CASE="$2"; shift 2 ;;
    --method) METHOD="$2"; shift 2 ;;
    --model) MODEL="$2"; shift 2 ;;
    --temperature) TEMPERATURE="$2"; shift 2 ;;
    --skip-ask) SKIP_ASK="true"; shift ;;
    *) echo "Unknown arg: $1" >&2; exit 1 ;;
  esac
done

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
bench_dir="$(dirname "$script_dir")"
repo_dir="$(dirname "$bench_dir")"

cd "$bench_dir"

if [[ ! -f ".env" ]]; then
  echo "[reproduce] Copying .env.example -> .env (remember to set OPENAI_API_KEY)."
  cp .env.example .env
fi

if [[ ! -d ".venv" ]]; then
  echo "[reproduce] Creating virtual environment."
  python -m venv .venv
fi

python_bin=".venv/bin/python"
if [[ ! -f "$python_bin" ]]; then
  echo "Python venv not found at $python_bin" >&2
  exit 1
fi

echo "[reproduce] Installing dependencies."
"$python_bin" -m pip install -e "$repo_dir"
"$python_bin" -m pip install -e .

echo "[reproduce] Extracting contexts."
"$python_bin" -m bench.cli extract --case "$CASE" --method "$METHOD"

if [[ "$SKIP_ASK" == "true" ]]; then
  echo "[reproduce] Skipping LLM inference."
else
  echo "[reproduce] Running LLM inference."
  "$python_bin" -m bench.cli ask --case "$CASE" --method "$METHOD" --model "$MODEL" --temperature "$TEMPERATURE"
fi

echo "[reproduce] Evaluating results."
"$python_bin" -m bench.cli eval --case "$CASE" --method "$METHOD"

echo "[reproduce] Generating reports."
"$python_bin" -m bench.cli report
"$python_bin" -m bench.cli report-public
