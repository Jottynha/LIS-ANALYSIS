from __future__ import annotations

import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.main import calcular_estatisticas_do_df, parse_lis_once


class LisAnalysisConsistencyTest(unittest.TestCase):
    SAMPLE_FILES = [
        Path("data/samples/Listas ATP/CONVENCIONAL/Caso0_ReEnergizacao_Convenc_SemControle.LIS"),
        Path("data/samples/Listas ATP/OTIMIZADA/caso0_Otimizada_SemControle.lis"),
    ]

    def test_parse_lis_once_extracts_table_and_time_series(self):
        for sample in self.SAMPLE_FILES:
            with self.subTest(sample=str(sample)):
                parsed = parse_lis_once(sample, verbose=False)

                self.assertIsNotNone(parsed.table_df, f"Tabela nao extraida de {sample}")
                self.assertFalse(parsed.table_df.empty, f"Tabela vazia em {sample}")
                self.assertIsNotNone(parsed.time_series_df, f"Series temporais nao extraidas de {sample}")
                self.assertFalse(parsed.time_series_df.empty, f"Series temporais vazias em {sample}")
                self.assertIn("mean", parsed.summary)
                self.assertIn("variance", parsed.summary)
                self.assertIn("std_dev", parsed.summary)
                self.assertIn("Step", parsed.time_series_df.columns)
                self.assertIn("Time", parsed.time_series_df.columns)

    def test_computed_grouped_stats_match_lis_summary(self):
        for sample in self.SAMPLE_FILES:
            with self.subTest(sample=str(sample)):
                parsed = parse_lis_once(sample, verbose=False)
                stats = calcular_estatisticas_do_df(parsed.table_df)
                summary = parsed.summary

                grouped_mean = float(summary["mean"][0])
                grouped_variance = float(summary["variance"][0])
                grouped_std_dev = float(summary["std_dev"][0])

                self.assertAlmostEqual(stats["mean"], grouped_mean, places=6)
                self.assertAlmostEqual(stats["variance"], grouped_variance, places=6)
                self.assertAlmostEqual(stats["std_dev"], grouped_std_dev, places=6)
                self.assertAlmostEqual(float(stats["bin_width"]), 0.05, places=9)
                self.assertGreater(float(stats["total_freq"]), 0.0)


if __name__ == "__main__":
    unittest.main()
