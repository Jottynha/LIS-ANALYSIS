from __future__ import annotations

import copy
import os
import re
import threading
from pathlib import Path
from typing import Any


class ATPParseError(RuntimeError):
    """Erro de parsing de arquivo ATP."""


_PARSE_CACHE_LOCK = threading.Lock()
_PARSE_CACHE_MAX_ITEMS = 8
_PARSE_CACHE: dict[str, dict[str, Any]] = {}


def _cache_path_key(path: Path) -> str:
    try:
        resolved = path.resolve()
    except Exception:
        resolved = path

    key = str(resolved)
    if os.name == "nt":
        key = key.lower()
    return key


def _file_signature(path: Path) -> tuple[int, int]:
    st = path.stat()
    return int(st.st_mtime_ns), int(st.st_size)


def invalidate_atp_parse_cache(path: str | Path | None = None) -> None:
    """Limpa cache do parser ATP (global ou por caminho específico)."""
    with _PARSE_CACHE_LOCK:
        if path is None:
            _PARSE_CACHE.clear()
            return

        key = _cache_path_key(Path(path))
        _PARSE_CACHE.pop(key, None)


def _parse_float(value: str) -> float:
    normalized = value.strip().replace("D", "E").replace("d", "e")
    return float(normalized)


def _extract_numeric_tokens(tokens: list[str]) -> list[float]:
    """Extrai tokens numéricos de uma lista de tokens, ignorando texto não numérico."""
    values: list[float] = []
    for tok in tokens:
        try:
            values.append(_parse_float(tok))
        except ValueError:
            continue
    return values


def _extract_numeric_fields(line: str) -> list[dict[str, Any]]:
    """Extrai campos numéricos com posição [start:end] no texto original da linha."""
    fields: list[dict[str, Any]] = []
    for match in re.finditer(r"\S+", line):
        token = match.group(0)
        try:
            value = _parse_float(token)
        except ValueError:
            continue
        fields.append(
            {
                "value": float(value),
                "start": match.start(),
                "end": match.end(),
                "text": token,
            }
        )
    return fields


def _fixed_width_field(
    raw_line: str,
    field: dict[str, Any],
    start: int,
    end: int,
) -> dict[str, Any]:
    """Usa a coluna fixa ATP quando o token detectado estiver dentro dela."""
    adjusted = dict(field)
    field_start = int(field["start"])
    field_end = int(field["end"])
    if start <= field_start and field_end <= end and end <= len(raw_line):
        adjusted["start"] = start
        adjusted["end"] = end
    return adjusted


def _mk_param(value: float, field: dict[str, Any], editable: bool = True) -> dict[str, Any]:
    return {
        "value": float(value),
        "original_value": float(value),
        "start": int(field["start"]),
        "end": int(field["end"]),
        "editable": bool(editable),
        "changed": False,
    }


def _parse_branch_line(
    tokens: list[str],
    raw_line: str,
    line_index: int,
) -> dict[str, Any] | None:
    """Parse de linha de /BRANCH preservando ordem dos valores numéricos (R, L, C opcional)."""
    numeric_fields = _extract_numeric_fields(raw_line)
    if not numeric_fields:
        return None

    if len(numeric_fields) >= 2:
        resistance = float(numeric_fields[0]["value"])
        inductance = float(numeric_fields[1]["value"])
    else:
        resistance = float(numeric_fields[0]["value"])
        inductance = None

    capacitance = None
    capacitance_is_control_default = False
    # Em arquivos ATPDraw, o terceiro valor em /BRANCH pode representar C,
    # inclusive quando for o valor padrão de controle (0).
    if len(numeric_fields) >= 3:
        c_candidate = float(numeric_fields[2]["value"])
        capacitance = c_candidate
        if abs(c_candidate) <= 0.0:
            capacitance_is_control_default = True

    params: dict[str, dict[str, Any]] = {
        "resistance": _mk_param(
            resistance,
            _fixed_width_field(raw_line, numeric_fields[0], 26, 32),
            editable=True,
        )
    }
    if inductance is not None:
        params["inductance"] = _mk_param(
            inductance,
            _fixed_width_field(raw_line, numeric_fields[1], 32, 38),
            editable=True,
        )
    if capacitance is not None and len(numeric_fields) >= 3:
        params["capacitance"] = _mk_param(
            capacitance,
            _fixed_width_field(raw_line, numeric_fields[2], 38, 44),
            editable=not capacitance_is_control_default,
        )

    return {
        "type": "branch",
        "resistance": resistance,
        "inductance": inductance,
        "capacitance": capacitance,
        "capacitance_is_control_default": capacitance_is_control_default,
        "parameters": params,
        "raw_line": raw_line,
        "line_index": line_index,
    }


def _parse_switch_line(tokens: list[str], raw_line: str, line_index: int) -> dict[str, Any] | None:
    """Parse de linha de /SWITCH: t_close e delay como dois primeiros numéricos da linha."""
    numeric_fields = _extract_numeric_fields(raw_line)
    if len(numeric_fields) < 2:
        return None

    t_close = float(numeric_fields[0]["value"])
    delay = float(numeric_fields[1]["value"])

    return {
        "type": "switch",
        "t_close": t_close,
        "delay": delay,
        "parameters": {
            "t_close": _mk_param(
                t_close,
                _fixed_width_field(raw_line, numeric_fields[0], 14, 24),
                editable=True,
            ),
            "delay": _mk_param(
                delay,
                _fixed_width_field(raw_line, numeric_fields[1], 24, 34),
                editable=True,
            ),
        },
        "raw_line": raw_line,
        "line_index": line_index,
    }


def _parse_source_line(tokens: list[str], raw_line: str, line_index: int) -> dict[str, Any] | None:
    """Parse de linha de /SOURCE: amplitude, frequência e fase opcional."""
    numeric_fields = _extract_numeric_fields(raw_line)
    if len(numeric_fields) < 2:
        return None

    amplitude = float(numeric_fields[0]["value"])
    frequency = float(numeric_fields[1]["value"])
    phase = float(numeric_fields[2]["value"]) if len(numeric_fields) >= 3 else None

    params: dict[str, dict[str, Any]] = {
        "amplitude": _mk_param(
            amplitude,
            _fixed_width_field(raw_line, numeric_fields[0], 10, 20),
            editable=True,
        ),
        "frequency": _mk_param(
            frequency,
            _fixed_width_field(raw_line, numeric_fields[1], 20, 30),
            editable=True,
        ),
    }
    if phase is not None:
        params["phase"] = _mk_param(
            phase,
            _fixed_width_field(raw_line, numeric_fields[2], 30, 40),
            editable=True,
        )

    return {
        "type": "source",
        "amplitude": amplitude,
        "frequency": frequency,
        "phase": phase,
        "parameters": params,
        "raw_line": raw_line,
        "line_index": line_index,
    }


def parse_atp_file(path: str | Path) -> list[dict[str, Any]]:
    """Lê um arquivo ATP (ATPDraw) por blocos e retorna elementos editáveis em formato dicionário."""
    atp_path = Path(path)
    if not atp_path.exists():
        raise FileNotFoundError(f"Arquivo ATP nao encontrado: {atp_path}")

    lines = atp_path.read_text(encoding="utf-8", errors="replace").splitlines()
    elements: list[dict[str, Any]] = []
    current_block = ""

    for idx, raw_line in enumerate(lines):
        line = raw_line.rstrip("\r\n")
        stripped = line.strip()

        if not stripped:
            continue

        upper_stripped = stripped.upper()
        if upper_stripped.startswith("/"):
            current_block = upper_stripped
            continue

        if stripped.startswith("C "):
            continue

        if current_block not in {"/BRANCH", "/SWITCH", "/SOURCE"}:
            continue

        tokens = stripped.split()
        try:
            parsed: dict[str, Any] | None
            if current_block == "/BRANCH":
                parsed = _parse_branch_line(
                    tokens,
                    raw_line,
                    idx,
                )
            elif current_block == "/SWITCH":
                parsed = _parse_switch_line(tokens, raw_line, idx)
            else:
                parsed = _parse_source_line(tokens, raw_line, idx)

            if parsed is not None:
                elements.append(parsed)
        except Exception:
            # Linha malformada é ignorada para manter robustez do parser.
            continue

    return elements


def parse_atp_file_cached(path: str | Path, force_refresh: bool = False) -> list[dict[str, Any]]:
    """Retorna parse ATP com cache por assinatura de arquivo (mtime_ns + tamanho)."""
    atp_path = Path(path)
    if not atp_path.exists():
        raise FileNotFoundError(f"Arquivo ATP nao encontrado: {atp_path}")

    signature = _file_signature(atp_path)
    cache_key = _cache_path_key(atp_path)

    with _PARSE_CACHE_LOCK:
        entry = _PARSE_CACHE.get(cache_key)
        if (
            not force_refresh
            and entry is not None
            and entry.get("signature") == signature
            and isinstance(entry.get("elements"), list)
        ):
            return copy.deepcopy(entry["elements"])

    elements = parse_atp_file(atp_path)

    with _PARSE_CACHE_LOCK:
        _PARSE_CACHE[cache_key] = {
            "signature": signature,
            "elements": copy.deepcopy(elements),
        }

        while len(_PARSE_CACHE) > _PARSE_CACHE_MAX_ITEMS:
            first_key = next(iter(_PARSE_CACHE))
            _PARSE_CACHE.pop(first_key, None)

    return elements


def get_editable_parameters(elements: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Converte elementos parseados em parâmetros editáveis para a GUI."""
    rows: list[dict[str, Any]] = []

    for idx, element in enumerate(elements):
        etype = str(element.get("type", "")).lower()

        if etype == "branch":
            if element.get("resistance") is not None:
                rows.append(
                    {
                        "label": "R (branch)",
                        "value": float(element["resistance"]),
                        "element_index": idx,
                        "field": "resistance",
                    }
                )
            if element.get("inductance") is not None:
                rows.append(
                    {
                        "label": "L (branch)",
                        "value": float(element["inductance"]),
                        "element_index": idx,
                        "field": "inductance",
                    }
                )
            if element.get("capacitance") is not None:
                is_default = bool(element.get("capacitance_is_control_default", False))
                rows.append(
                    {
                        "label": "C (branch) - valor padrao (0)" if is_default else "C (branch)",
                        "value": float(element["capacitance"]),
                        "element_index": idx,
                        "field": "capacitance",
                        "editable": not is_default,
                    }
                )

        elif etype == "switch":
            if element.get("t_close") is not None:
                rows.append(
                    {
                        "label": "Switch time",
                        "value": float(element["t_close"]),
                        "element_index": idx,
                        "field": "t_close",
                    }
                )
            if element.get("delay") is not None:
                rows.append(
                    {
                        "label": "Switch delay",
                        "value": float(element["delay"]),
                        "element_index": idx,
                        "field": "delay",
                    }
                )

        elif etype == "source":
            if element.get("amplitude") is not None:
                rows.append(
                    {
                        "label": "Amplitude",
                        "value": float(element["amplitude"]),
                        "element_index": idx,
                        "field": "amplitude",
                    }
                )
            if element.get("frequency") is not None:
                rows.append(
                    {
                        "label": "Frequency",
                        "value": float(element["frequency"]),
                        "element_index": idx,
                        "field": "frequency",
                    }
                )
            if element.get("phase") is not None:
                rows.append(
                    {
                        "label": "Phase",
                        "value": float(element["phase"]),
                        "element_index": idx,
                        "field": "phase",
                    }
                )

    return rows


def update_parameter(
    elements: list[dict[str, Any]],
    element_name: str,
    new_value: float | str,
    line_index: int | None = None,
    parameter_name: str | None = None,
) -> dict[str, Any]:
    """Atualiza campo numérico de elemento parseado por line_index + campo."""
    if parameter_name is None:
        raise ATPParseError("parameter_name e obrigatorio para atualizar elemento por bloco")

    try:
        parsed_value = _parse_float(str(new_value))
    except ValueError as exc:
        raise ATPParseError(f"Valor numerico invalido: {new_value}") from exc

    if line_index is None:
        raise ATPParseError("line_index e obrigatorio para atualizar elemento por bloco")

    match = None
    for element in elements:
        if int(element.get("line_index", -1)) == int(line_index):
            match = element
            break

    if match is None:
        raise ATPParseError(f"Elemento nao encontrado para line_index={line_index}")

    if parameter_name not in match:
        raise ATPParseError(
            f"Campo '{parameter_name}' nao existe no elemento da linha {line_index}"
        )

    match[parameter_name] = float(parsed_value)
    params = match.get("parameters")
    if isinstance(params, dict) and parameter_name in params:
        params[parameter_name]["value"] = float(parsed_value)
        params[parameter_name]["changed"] = abs(
            float(parsed_value) - float(params[parameter_name].get("original_value", parsed_value))
        ) > 1e-15
    return match
