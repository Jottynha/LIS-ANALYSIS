from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.atp_parser import parse_atp_file
from lis_analysis.atp_writer import assert_identity_when_no_overrides, write_atp_file


class ATPWriterIdentityTest(unittest.TestCase):
    def test_roundtrip_without_changes_is_byte_identical(self):
        sample_path = Path("data/samples/ACP/caso1_convenc_RPI.atp")
        if not sample_path.exists():
            self.skipTest("Arquivo de referência não encontrado para teste de identidade")

        original_bytes = sample_path.read_bytes()
        with sample_path.open("r", encoding="latin-1", errors="replace", newline="") as f:
            original_lines = f.read().splitlines(keepends=True)
        elements = parse_atp_file(sample_path)

        # Validação em memória (sem escrita em disco)
        assert_identity_when_no_overrides(elements, original_lines)

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / sample_path.name
            write_atp_file(elements, original_lines, out_path)
            self.assertEqual(original_bytes, out_path.read_bytes())


if __name__ == "__main__":
    unittest.main()
