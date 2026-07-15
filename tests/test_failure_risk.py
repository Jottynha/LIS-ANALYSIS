from __future__ import annotations

import math
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.failure_risk import FailureRiskConfig, calculate_failure_risk


class FailureRiskArticleReferenceTest(unittest.TestCase):
    def test_reproduces_article_table_5_case_zero(self):
        # Artigo: V50=1.68 p.u., sigma_s=15.22%, Vbase=429 kV,
        # 201 torres x 4 cadeias = 804 gaps em paralelo.
        mean_pu = 1.68
        std_pu = mean_pu * 0.1522
        references = [
            (3.50, 1041.0, 1.00e-3),
            (3.00, 939.0, 1.27e-2),
            (2.50, 827.0, 8.75e-2),
        ]

        for distance, expected_cfo, expected_risk in references:
            with self.subTest(distance=distance):
                result = calculate_failure_risk(
                    mean_pu,
                    std_pu,
                    FailureRiskConfig(
                        conductor_structure_distance_m=distance,
                        insulation_distance_m=distance,
                    ),
                )
                self.assertAlmostEqual(result.corrected_cfo_kv, expected_cfo, delta=7.0)
                self.assertAlmostEqual(result.risk, expected_risk, delta=expected_risk * 0.15)

    def test_converts_lis_statistics_from_pu_to_kv(self):
        result = calculate_failure_risk(1.68, 0.20, FailureRiskConfig())
        self.assertAlmostEqual(result.mean_overvoltage_kv, 1.68 * 429.0)
        self.assertAlmostEqual(result.switching_std_kv, 0.20 * 429.0)

    def test_uses_fifth_root_of_parallel_gaps(self):
        result = calculate_failure_risk(
            1.0,
            0.1,
            FailureRiskConfig(parallel_gaps=32.0),
        )
        self.assertAlmostEqual(
            result.corrected_withstand_std_kv,
            result.withstand_std_kv / 2.0,
        )

    def test_rejects_non_finite_or_non_physical_inputs(self):
        with self.assertRaises(ValueError):
            calculate_failure_risk(math.nan, 0.1, FailureRiskConfig())
        with self.assertRaises(ValueError):
            calculate_failure_risk(
                1.0,
                0.1,
                FailureRiskConfig(conductor_structure_distance_m=0.0),
            )


if __name__ == "__main__":
    unittest.main()
