from __future__ import annotations

from pathlib import Path
from typing import Any


def _format_value(value: float) -> str:
    return f"{float(value):.10g}".upper()


def _leading_whitespace(text: str) -> str:
    return text[: len(text) - len(text.lstrip(" \t"))]


def format_element_line(element: dict[str, Any], newline: str = "\n") -> str:
    """Formata linha ATP por tipo de bloco usando colunas fixas simples."""
    etype = str(element.get("type", "")).lower()
    raw_line = str(element.get("raw_line", ""))
    stripped_tokens = raw_line.strip().split()
    if not stripped_tokens:
        return raw_line + ("" if raw_line.endswith(("\n", "\r")) else newline)

    indent = _leading_whitespace(raw_line)
    head = stripped_tokens[0]

    if etype == "branch":
        r = _format_value(float(element["resistance"]))
        l = element.get("inductance")
        c = element.get("capacitance")
        if l is None:
            if c is None:
                return f"{indent}{head:<22}{r:>12}".rstrip() + newline
            c_fmt = _format_value(float(c))
            return f"{indent}{head:<22}{r:>12}{c_fmt:>12}".rstrip() + newline
        l_fmt = _format_value(float(l))
        if c is None:
            return f"{indent}{head:<22}{r:>12}{l_fmt:>12}".rstrip() + newline
        c_fmt = _format_value(float(c))
        return f"{indent}{head:<22}{r:>12}{l_fmt:>12}{c_fmt:>12}".rstrip() + newline

    if etype == "switch":
        t_close = _format_value(float(element["t_close"]))
        delay = _format_value(float(element["delay"]))
        return f"{indent}{head:<22}{t_close:>12}{delay:>12}".rstrip() + newline

    if etype == "source":
        amp = _format_value(float(element["amplitude"]))
        freq = _format_value(float(element["frequency"]))
        phase = element.get("phase")
        if phase is None:
            return f"{indent}{head:<10}{amp:>16}{freq:>12}".rstrip() + newline
        phase_fmt = _format_value(float(phase))
        return f"{indent}{head:<10}{amp:>16}{freq:>12}{phase_fmt:>12}".rstrip() + newline

    return raw_line + ("" if raw_line.endswith(("\n", "\r")) else newline)


def write_atp_file(
    elements: list[dict[str, Any]],
    original_lines: list[str],
    output_path: str | Path,
) -> Path:
    """Escreve novo ATP preservando linhas desconhecidas e reescrevendo linhas editáveis."""
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    by_index = {int(e.get("line_index", -1)): e for e in elements if "line_index" in e}
    rendered: list[str] = []

    for idx, line in enumerate(original_lines):
        element = by_index.get(idx)
        if element is None:
            rendered.append(line)
            continue

        newline = "\n"
        if line.endswith("\r\n"):
            newline = "\r\n"
        elif line.endswith("\n"):
            newline = "\n"
        elif line:
            newline = ""

        try:
            rendered.append(format_element_line(element, newline=newline))
        except Exception:
            rendered.append(line)

    out_path.write_text("".join(rendered), encoding="utf-8", errors="replace")
    return out_path
