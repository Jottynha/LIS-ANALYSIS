from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path
import sys
import importlib

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

_atp_parser = importlib.import_module("lis_analysis.atp_parser")
invalidate_atp_parse_cache = _atp_parser.invalidate_atp_parse_cache
parse_atp_file_cached = _atp_parser.parse_atp_file_cached


def _write_branch_file(path: Path, resistance: str) -> None:
    content = (
        "BEGIN NEW DATA CASE\n"
        "/BRANCH\n"
        f"  X0001AX0003A                {resistance}   75.                                         0\n"
        "BLANK BRANCH\n"
    )
    path.write_text(content, encoding="utf-8")


class ATPParserCacheTest(unittest.TestCase):
    def test_cached_result_isolated_from_mutations(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            atp_path = Path(tmpdir) / "case.atp"
            _write_branch_file(atp_path, "5.")

            invalidate_atp_parse_cache(atp_path)
            first = parse_atp_file_cached(atp_path)
            self.assertTrue(first)

            first[0]["resistance"] = 999.0
            first[0]["parameters"]["resistance"]["value"] = 999.0

            second = parse_atp_file_cached(atp_path)
            self.assertEqual(float(second[0]["resistance"]), 5.0)
            self.assertEqual(float(second[0]["parameters"]["resistance"]["value"]), 5.0)

    def test_cache_invalidation_when_file_changes(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            atp_path = Path(tmpdir) / "case.atp"
            _write_branch_file(atp_path, "5.")

            invalidate_atp_parse_cache(atp_path)
            initial = parse_atp_file_cached(atp_path)
            self.assertEqual(float(initial[0]["resistance"]), 5.0)

            time.sleep(0.01)
            _write_branch_file(atp_path, "9.")

            updated = parse_atp_file_cached(atp_path)
            self.assertEqual(float(updated[0]["resistance"]), 9.0)


if __name__ == "__main__":
    unittest.main()
