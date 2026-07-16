from pathlib import Path
import queue
import threading
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

def test_worker_callback_runs_only_when_main_thread_drains_queue():
    app = ModernLisAnalysisApp.__new__(ModernLisAnalysisApp)
    app._ui_thread_id = threading.get_ident()
    app._ui_task_queue = queue.Queue()
    app._ui_queue_closed = False
    app.after = lambda *_args, **_kwargs: None
    calls = []

    worker = threading.Thread(
        target=app._run_on_ui_thread,
        args=(calls.append, "executado"),
    )
    worker.start()
    worker.join()

    assert calls == []
    assert app._ui_task_queue.qsize() == 1

    app._drain_ui_tasks()

    assert calls == ["executado"]
    assert app._ui_task_queue.empty()

def test_background_workers_do_not_access_tk_objects_directly():
    for method_name in (
        "_run_atp_parameter_sweep",
        "_run_atp_simulation",
        "_process_selected",
    ):
        worker = _nested_code(getattr(ModernLisAnalysisApp, method_name), "worker")
        forbidden = {
            name
            for name in worker.co_names
            if name == "after" or name == "log_textbox" or name.endswith("_var")
        }

        assert forbidden == set(), (method_name, forbidden)

def test_batch_postprocessing_never_opens_generated_plots():
    method = ModernLisAnalysisApp._postprocess_atp_sweep_lis

    assert "_open_file_in_editor" not in method.__code__.co_names
    assert "show_plots" not in method.__code__.co_varnames
