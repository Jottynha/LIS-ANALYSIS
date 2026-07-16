from pathlib import Path
from types import CodeType
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.gui import ModernLisAnalysisApp


def _nested_code(function, name: str) -> CodeType:
    return next(
        constant
        for constant in function.__code__.co_consts
        if isinstance(constant, CodeType) and constant.co_name == name
    )


def test_single_simulation_worker_captures_selected_atp_executable():
    worker = _nested_code(ModernLisAnalysisApp._run_atp_simulation, "worker")

    assert "atp_executable_path" in worker.co_freevars


def test_export_parameters_does_not_require_atp_executable():
    export_function = ModernLisAnalysisApp._export_atp_parameters_txt

    assert "_resolve_atp_executable_for_run" not in export_function.__code__.co_names
