from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.gui import ModernLisAnalysisApp


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


def test_separate_tower_and_insulator_counts_produce_parallel_gaps():
    app = ModernLisAnalysisApp.__new__(ModernLisAnalysisApp)
    app.enable_risk_var = FakeVar(True)
    app.risk_base_voltage_var = FakeVar("429")
    app.risk_height_var = FakeVar("17.98")
    app.risk_distance_var = FakeVar("3.50")
    app.risk_width_var = FakeVar("2")
    app.risk_subconductors_var = FakeVar("4")
    app.risk_insulation_distance_var = FakeVar("3.50")
    app.risk_tower_count_var = FakeVar("201")
    app.risk_insulator_chains_per_tower_var = FakeVar("4")

    config = app._collect_risk_config()

    assert config.parallel_gaps == 804


def test_legacy_parallel_gaps_are_split_without_changing_total():
    assert ModernLisAnalysisApp._split_legacy_parallel_gaps("804") == ("201", "4")
    assert ModernLisAnalysisApp._split_legacy_parallel_gaps("801") == ("801", "1")
