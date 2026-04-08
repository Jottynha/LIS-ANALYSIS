#!/usr/bin/env bash
set -euo pipefail

# Script de apoio para desrastrear artefatos gerados sem apagar arquivos locais.
# Uso:
#   ./scripts/cleanup_git_artifacts.sh --dry-run
#   ./scripts/cleanup_git_artifacts.sh --apply

MODE="dry-run"

if [[ "${1:-}" == "--apply" ]]; then
    MODE="apply"
elif [[ "${1:-}" == "--dry-run" || -z "${1:-}" ]]; then
    MODE="dry-run"
else
    echo "Uso: $0 [--dry-run|--apply]"
    exit 2
fi

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"

PATTERN='(^Simulation_Result/|(^|/)__pycache__/|\.pyc$|_parametros_detectados\.txt$|_param(_[0-9]+)?\.atp$|__param_[0-9]{8}_[0-9]{6}_[0-9]+\.atp$)'

matches=()
while IFS= read -r path; do
    [[ -n "$path" ]] && matches+=("$path")
done < <(git ls-files | grep -E "$PATTERN" || true)

if [[ ${#matches[@]} -eq 0 ]]; then
    echo "Nenhum artefato rastreado encontrado para limpeza."
    exit 0
fi

echo "Artefatos encontrados: ${#matches[@]}"
printf ' - %s\n' "${matches[@]}"

if [[ "$MODE" == "dry-run" ]]; then
    echo
    echo "Dry-run: nada foi alterado."
    echo "Para aplicar: ./scripts/cleanup_git_artifacts.sh --apply"
    exit 0
fi

for f in "${matches[@]}"; do
    git rm --cached -- "$f"
done

echo
echo "Limpeza concluida. Verifique com: git status --short"
