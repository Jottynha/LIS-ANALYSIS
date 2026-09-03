from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.gui import ModernLisAnalysisApp


class FakeVar:
    def __init__(self, value: str):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def _build_app():
    app = ModernLisAnalysisApp.__new__(ModernLisAnalysisApp)
    app._atp_param_rows = [
        {
            "line_index": 10,
            "parameter": "resistance",
            "var": FakeVar("5.0"),
            "_updating": False,
        }
    ]
    app.parameter_overrides = {}
    app._atp_parameter_editor_snapshot = {}
    app._atp_parameter_editor_overrides_snapshot = {}
    return app


def test_parameter_editor_detects_unsaved_changes():
    app = _build_app()
    app._capture_atp_parameter_editor_snapshot()

    assert not app._atp_parameter_editor_has_unsaved_changes()

    app._atp_param_rows[0]["var"].set("10.0")

    assert app._atp_parameter_editor_has_unsaved_changes()


def test_parameter_editor_discard_restores_opening_state():
    app = _build_app()
    app.parameter_overrides = {(10, "resistance"): 5.0}
    app._capture_atp_parameter_editor_snapshot()
    app._atp_param_rows[0]["var"].set("20.0")
    app.parameter_overrides = {(10, "resistance"): 20.0}
    app._update_atp_parameter_row = lambda *_args, **_kwargs: None
    app._refresh_atp_param_status = lambda: None
    app._apply_atp_parameter_filter = lambda: None

    app._restore_atp_parameter_editor_snapshot()

    assert app._atp_param_rows[0]["var"].get() == "5.0"
    assert app.parameter_overrides == {(10, "resistance"): 5.0}
