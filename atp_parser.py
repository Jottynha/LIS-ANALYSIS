from __future__ import annotations

from pathlib import Path
from typing import Any

from atp_elements import ATPElement, element_class_for

SUPPORTED_TYPES = {"R", "L", "C", "V", "I"}


class ATPParseError(RuntimeError):
    """Erro de parsing de arquivo ATP."""


def _is_continuation_line(line: str) -> bool:
    if not line:
        return False
    return line.startswith(" ") or line.startswith("+")


def _parse_float(value: str) -> float:
    normalized = value.strip().replace("D", "E").replace("d", "e")
    return float(normalized)


def _parse_card_tokens(card_text: str) -> tuple[str, str, str, str, list[str]] | None:
    tokens = card_text.split()
    if len(tokens) < 4:
        return None

    name = tokens[0].strip()
    if not name:
        return None

    card_type = name[0].upper()
    if card_type not in SUPPORTED_TYPES:
        return None

    node1 = tokens[1].strip()
    node2 = tokens[2].strip()
    value_token = tokens[3].strip()

    if not node1 or not node2 or not value_token:
        return None

    try:
        _parse_float(value_token)
    except ValueError:
        return None

    tail = tokens[4:] if len(tokens) > 4 else []
    return card_type, name, node1, node2, [value_token] + tail


def _card_value_parameter_name(card_type: str) -> str:
    mapping = {
        "R": "resistance",
        "L": "inductance",
        "C": "capacitance",
        "V": "voltage",
        "I": "current",
    }
    return mapping.get(card_type.upper(), "value")


def parse_atp_file(path: str | Path) -> tuple[list[ATPElement], list[str]]:
    """Lê um arquivo ATP e retorna elementos editáveis e linhas originais."""
    atp_path = Path(path)
    if not atp_path.exists():
        raise FileNotFoundError(f"Arquivo ATP nao encontrado: {atp_path}")

    original_lines = atp_path.read_text(encoding="utf-8", errors="replace").splitlines(keepends=True)

    logical_cards: list[dict[str, Any]] = []
    current_lines: list[str] = []
    current_start = 0

    def flush_current(end_index: int) -> None:
        nonlocal current_lines, current_start
        if not current_lines:
            return

        logical_cards.append(
            {
                "start": current_start,
                "end": end_index,
                "lines": current_lines[:],
            }
        )
        current_lines = []

    for idx, line in enumerate(original_lines):
        if not current_lines:
            current_start = idx
            current_lines.append(line)
            continue

        if _is_continuation_line(line):
            current_lines.append(line)
            continue

        flush_current(idx - 1)
        current_start = idx
        current_lines = [line]

    flush_current(len(original_lines) - 1)

    elements: list[ATPElement] = []

    for card in logical_cards:
        first_line = card["lines"][0]
        if not first_line.strip():
            continue

        first_char = first_line[0].upper()
        if first_char not in SUPPORTED_TYPES:
            continue

        joined_parts = []
        for raw in card["lines"]:
            no_nl = raw.rstrip("\r\n")
            if not joined_parts:
                joined_parts.append(no_nl)
            else:
                joined_parts.append(no_nl.lstrip(" +"))
        joined_text = " ".join(joined_parts)

        parsed = _parse_card_tokens(joined_text)
        if parsed is None:
            continue

        card_type, name, node1, node2, value_and_tail = parsed
        value_token = value_and_tail[0]
        tail_tokens = value_and_tail[1:]

        parameter_name = _card_value_parameter_name(card_type)
        value = _parse_float(value_token)

        element_cls = element_class_for(card_type)
        element = element_cls(
            type=card_type,
            name=name,
            nodes=[node1, node2],
            parameters={
                parameter_name: value,
                "_tail": " ".join(tail_tokens),
                "_value_token": value_token,
            },
            raw_line=joined_text,
            line_index=card["start"],
            line_end_index=card["end"],
            continuation_lines=card["lines"][1:],
            modified=False,
        )
        elements.append(element)

    return elements, original_lines


def get_editable_parameters(elements: list[ATPElement]) -> list[dict[str, Any]]:
    """Retorna estrutura amigável para UI com parâmetros editáveis."""
    rows: list[dict[str, Any]] = []
    for element in elements:
        parameter_name = element.parameter_name()
        rows.append(
            {
                "line_index": element.line_index,
                "name": element.name,
                "type": element.__class__.__name__,
                "parameter": parameter_name,
                "value": element.parameters.get(parameter_name),
            }
        )
    return rows


def update_parameter(
    elements: list[ATPElement],
    element_name: str,
    new_value: float | str,
    line_index: int | None = None,
    parameter_name: str | None = None,
) -> ATPElement:
    """Atualiza parâmetro de um elemento por nome (e opcionalmente linha/parâmetro)."""
    matches = [e for e in elements if e.name.lower() == element_name.lower()]
    if line_index is not None:
        matches = [e for e in matches if e.line_index == line_index]

    if parameter_name is not None:
        matches = [e for e in matches if e.parameter_name().lower() == parameter_name.lower()]

    if not matches:
        raise ATPParseError(f"Elemento nao encontrado para atualizacao: {element_name}")

    if len(matches) > 1:
        raise ATPParseError(
            f"Elemento ambiguo '{element_name}'. Informe line_index para diferenciar ocorrencias."
        )

    element = matches[0]
    try:
        parsed_value = _parse_float(str(new_value))
    except ValueError as exc:
        raise ATPParseError(f"Valor numerico invalido para {element.name}: {new_value}") from exc

    element.set_value(parsed_value)
    return element
