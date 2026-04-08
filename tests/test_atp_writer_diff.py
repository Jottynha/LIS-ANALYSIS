from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import sys
import importlib

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

_atp_parser = importlib.import_module("lis_analysis.atp_parser")
_atp_writer = importlib.import_module("lis_analysis.atp_writer")

parse_atp_file = _atp_parser.parse_atp_file
update_parameter = _atp_parser.update_parameter
write_atp_file = _atp_writer.write_atp_file


def get_diff_ranges(line1: str, line2: str):
    diffs = []
    current = None

    for i, (c1, c2) in enumerate(zip(line1, line2)):
        if c1 != c2:
            if current is None:
                current = [i, i]
            else:
                current[1] = i
        else:
            if current:
                diffs.append(tuple(current))
                current = None

    if current:
        diffs.append(tuple(current))

    return diffs


def _read_lines(path: Path) -> list[str]:
    return path.read_text(encoding="latin-1", errors="replace").splitlines()


def _assert_single_small_diff(original_lines: list[str], modified_lines: list[str]) -> None:
    assert len(original_lines) == len(modified_lines)

    changed_lines = []
    for i, (l1, l2) in enumerate(zip(original_lines, modified_lines)):
        if l1 != l2:
            changed_lines.append((i, l1, l2))

    # Deve alterar apenas 1 linha
    assert len(changed_lines) == 1, f"Esperado 1 linha alterada, mas encontrou {len(changed_lines)}"

    idx, l1, l2 = changed_lines[0]

    # Mesmo tamanho
    assert len(l1) == len(l2), "Linha alterada mudou de tamanho (ERRO GRAVE)"

    diffs = get_diff_ranges(l1, l2)

    # Deve haver apenas 1 bloco de diferença
    assert len(diffs) == 1, f"Múltiplas regiões alteradas: {diffs}"

    start, end = diffs[0]

    # Diferença deve ser pequena (campo numérico)
    assert (end - start) < 10, f"Alteração muito grande: {start}-{end}"

    # Extra: o restante da linha precisa estar intacto
    assert l1[:start] == l2[:start], "Prefixo da linha foi alterado fora do campo esperado"
    assert l1[end + 1 :] == l2[end + 1 :], "Sufixo da linha foi alterado fora do campo esperado"

    # Extra: valor numérico realmente mudou no bloco alterado
    old_num = float(l1[start : end + 1].strip())
    new_num = float(l2[start : end + 1].strip())
    assert old_num != new_num, "Valor numérico do campo alterado não mudou"

    print(f"Linha alterada: {idx}")
    print(f"Intervalo alterado: {start}-{end}")


def test_single_parameter_change():
    base_dir = Path(__file__).parent / "data"
    original_lines = _read_lines(base_dir / "original.atp")
    modified_lines = _read_lines(base_dir / "modificado.atp")
    _assert_single_small_diff(original_lines, modified_lines)


class ATPWriterDiffTest(unittest.TestCase):
    def test_generated_file_single_parameter_change(self):
        base_dir = Path(__file__).parent / "data"
        original_path = base_dir / "original.atp"

        with original_path.open("r", encoding="latin-1", errors="replace", newline="") as f:
            original_lines_keepends = f.read().splitlines(keepends=True)

        elements = parse_atp_file(original_path)
        target_element = next(
            e for e in elements if e.get("type") == "branch" and e.get("resistance") is not None
        )
        line_index = int(target_element["line_index"])

        update_parameter(
            elements,
            element_name="branch",
            new_value=7.0,
            line_index=line_index,
            parameter_name="resistance",
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            generated_path = Path(tmpdir) / "modificado.atp"
            write_atp_file(elements, original_lines_keepends, generated_path)

            original_lines = _read_lines(original_path)
            generated_lines = _read_lines(generated_path)

            _assert_single_small_diff(original_lines, generated_lines)

            changed = [(i, a, b) for i, (a, b) in enumerate(zip(original_lines, generated_lines)) if a != b]
            changed_idx, old_line, new_line = changed[0]
            self.assertEqual(changed_idx, line_index)
            diff_ranges = get_diff_ranges(old_line, new_line)
            diff_start, diff_end = diff_ranges[0]
            old_field = old_line[diff_start : diff_end + 1].strip()
            new_field = new_line[diff_start : diff_end + 1].strip()

            # Preserva estilo ATP para inteiros com ponto final (ex.: 5. -> 7.)
            if old_field.endswith("."):
                self.assertTrue(
                    new_field.endswith("."),
                    f"Formato com ponto final nao preservado: {old_field!r} -> {new_field!r}",
                )

            generated_elements = parse_atp_file(generated_path)
            generated_target = next(
                e for e in generated_elements if e.get("type") == "branch" and e.get("line_index") == line_index
            )
            self.assertAlmostEqual(float(generated_target["resistance"]), 7.0)


if __name__ == "__main__":
    unittest.main()
