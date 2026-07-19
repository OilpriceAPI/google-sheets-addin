#!/usr/bin/env bash
set -euo pipefail

patterns=(
  '[0-9a-fA-F]{64}'
  'sk-[A-Za-z0-9_-]{20,}'
  'AIza[A-Za-z0-9_-]{30,}'
  'gh[pousr]_[A-Za-z0-9_]{20,}'
  '(?i)(api[_-]?key|access[_-]?token|secret)[^[:cntrl:]]{0,40}[=:][[:space:]]*[A-Za-z0-9_-]{32,}'
)

found=''
for pattern in "${patterns[@]}"; do
  matches="$(
    rg -l --hidden \
      --glob '!.git/**' \
      --glob '!node_modules/**' \
      --glob '!scripts/scan-secrets.sh' \
      --glob '!package-lock.json' \
      "$pattern" . || true
  )"
  if [[ -n "$matches" ]]; then
    found+="$matches"$'\n'
  fi
done

if [[ -n "$found" ]]; then
  printf 'Potential secret pattern found in:\n'
  printf '%s' "$found" | sort -u
  exit 1
fi

printf 'Secret scan passed (filenames only).\n'
