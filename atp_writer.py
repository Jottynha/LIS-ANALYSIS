from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


def _fit_value_to_width(new_value: float, width: int) -> tuple[str, bool]:
    """Formata valor para caber em largura fixa, alinhado à direita."""
    if width <= 0:
        return "", False

    candidates = [
        f"{float(new_value):.10g}",
        f"{float(new_value):.8g}",
        f"{float(new_value):.6g}",
        f"{float(new_value):.4g}",
        f"{float(new_value):.3e}",
    ]

    for c in candidates:
        if len(c) <= width:
            return c.rjust(width), False

    # Truncamento de segurança: nunca expande a linha.
    raw = candidates[-1]
    return raw[-width:].rjust(width), True


def replace_value_in_line(line: str, start: int, end: int, new_value: float) -> str:
    """Substitui apenas o trecho [start:end], mantendo tamanho/alinhamento do campo."""
    if start < 0 or end < start or end > len(line):
        return line

    field = line[start:end]
    width = len(field)
    formatted, truncated = _fit_value_to_width(new_value, width)
    if truncated:
        logger.warning(
            "Valor %.12g truncado para caber no campo [%s:%s] largura=%s",
            float(new_value),
            start,
            end,
            width,
        )

    return line[:start] + formatted + line[end:]


def apply_parameter_overrides(lines: list[str], overrides: dict[int, list[dict[str, Any]]]) -> list[str]:
    """Aplica alterações no .atp sem quebrar alinhamento, editando por posição."""
    new_lines = list(lines)

    for line_index, changes in overrides.items():
        if line_index < 0 or line_index >= len(new_lines):
            continue

        original_line = new_lines[line_index]
        updated_line = original_line

        # Direita -> esquerda para não deslocar índices.
        ordered = sorted(changes, key=lambda c: int(c.get("start", -1)), reverse=True)
        for change in ordered:
            if not bool(change.get("editable", True)):
                continue

            start = int(change.get("start", -1))
            end = int(change.get("end", -1))
            old_value = change.get("old_value")
            new_value = float(change.get("new_value"))
            field_name = str(change.get("field", "value"))

            updated_line = replace_value_in_line(updated_line, start, end, new_value)
            logger.info(
                "Linha %s: %s alterado de %s para %s",
                line_index + 1,
                field_name,
                old_value,
                new_value,
            )

        new_lines[line_index] = updated_line

    return new_lines


def write_atp_file(
    elements: list[dict[str, Any]],
    original_lines: list[str],
    output_path: str | Path,
) -> Path:
    """Escreve novo ATP aplicando somente alterações pontuais de parâmetros por posição."""
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    overrides_by_line: dict[int, list[dict[str, Any]]] = {}
    for element in elements:
        line_index = int(element.get("line_index", -1))
        params = element.get("parameters", {})
        if not isinstance(params, dict):
            continue

        for field_name, meta in params.items():
            if not isinstance(meta, dict):
                continue
            if not bool(meta.get("changed", False)):
                continue
            if not bool(meta.get("editable", True)):
                continue

            change = {
                "field": field_name,
                "start": int(meta.get("start", -1)),
                "end": int(meta.get("end", -1)),
                "old_value": meta.get("original_value", meta.get("value")),
                "new_value": float(meta.get("value")),
                "editable": bool(meta.get("editable", True)),
            }
            overrides_by_line.setdefault(line_index, []).append(change)

    rendered_lines = apply_parameter_overrides(original_lines, overrides_by_line)
    out_path.write_text("".join(rendered_lines), encoding="latin-1", errors="replace")
    return out_path


def assert_identity_when_no_overrides(elements: list[dict[str, Any]], original_lines: list[str]) -> None:
    """Valida que não há diferenças byte a byte quando nenhum parâmetro foi alterado."""
    overrides_by_line: dict[int, list[dict[str, Any]]] = {}
    for element in elements:
        line_index = int(element.get("line_index", -1))
        params = element.get("parameters", {})
        if not isinstance(params, dict):
            continue
        for field_name, meta in params.items():
            if not isinstance(meta, dict):
                continue
            if not bool(meta.get("changed", False)):
                continue
            if not bool(meta.get("editable", True)):
                continue
            change = {
                "field": field_name,
                "start": int(meta.get("start", -1)),
                "end": int(meta.get("end", -1)),
                "old_value": meta.get("original_value", meta.get("value")),
                "new_value": float(meta.get("value")),
                "editable": bool(meta.get("editable", True)),
            }
            overrides_by_line.setdefault(line_index, []).append(change)

    new_lines = apply_parameter_overrides(original_lines, overrides_by_line)
    if b"".join(s.encode("latin-1", errors="replace") for s in new_lines) != b"".join(
        s.encode("latin-1", errors="replace") for s in original_lines
    ):
        raise AssertionError("Roundtrip sem alterações não preservou bytes do arquivo ATP")
