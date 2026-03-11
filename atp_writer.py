from __future__ import annotations

from pathlib import Path

from atp_elements import ATPElement


def _format_value(value: float) -> str:
    return f"{float(value):.6g}".upper()


def format_element_line(element: ATPElement, newline: str = "\n") -> str:
    """Formata uma linha de componente ATP em colunas fixas."""
    if len(element.nodes) < 2:
        raise ValueError(f"Elemento {element.name} sem nos suficientes para escrita")

    name = element.name[:6]
    node1 = element.nodes[0][:8]
    node2 = element.nodes[1][:8]
    value = _format_value(element.get_value())
    tail = str(element.parameters.get("_tail", "")).strip()

    base = f"{name:<6}{node1:<8}{node2:<8}{value:>12}"
    if tail:
        base = f"{base} {tail}"
    return base.rstrip() + newline


def write_atp_file(
    elements: list[ATPElement],
    original_lines: list[str],
    output_path: str | Path,
) -> Path:
    """Escreve novo ATP preservando linhas desconhecidas e alterando apenas componentes modificados."""
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    by_start = {e.line_index: e for e in elements}
    consumed_until = -1
    rendered: list[str] = []

    for idx, line in enumerate(original_lines):
        if idx <= consumed_until:
            continue

        element = by_start.get(idx)
        if element is None:
            rendered.append(line)
            continue

        line_end = max(element.line_end_index, element.line_index)
        consumed_until = line_end

        if not element.modified:
            rendered.extend(original_lines[element.line_index : line_end + 1])
            continue

        newline = "\n"
        first_original = original_lines[element.line_index] if element.line_index < len(original_lines) else ""
        if first_original.endswith("\r\n"):
            newline = "\r\n"
        elif first_original.endswith("\n"):
            newline = "\n"
        elif first_original:
            newline = ""

        rendered.append(format_element_line(element, newline=newline))

        # Mantém linhas de continuação originais para evitar perda estrutural não suportada.
        if line_end > element.line_index:
            rendered.extend(original_lines[element.line_index + 1 : line_end + 1])

    out_path.write_text("".join(rendered), encoding="utf-8", errors="replace")
    return out_path
