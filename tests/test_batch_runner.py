from __future__ import annotations

import tempfile
import threading
import time
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.atp_parser import parse_atp_file
from lis_analysis.batch_runner import (
    SweepParameterRef,
    generate_sweep_values,
    run_parameter_sweep,
)


class BatchRunnerSweepValuesTest(unittest.TestCase):
    def test_generate_values_ascending_inclusive(self):
        self.assertEqual(generate_sweep_values(5, 20, 5), [5.0, 10.0, 15.0, 20.0])

    def test_generate_values_descending_inclusive(self):
        self.assertEqual(generate_sweep_values(20, 5, -5), [20.0, 15.0, 10.0, 5.0])

    def test_generate_values_single_value_when_start_equals_stop(self):
        self.assertEqual(generate_sweep_values(3, 3, 1), [3.0])

    def test_generate_values_rejects_zero_step(self):
        with self.assertRaises(ValueError):
            generate_sweep_values(1, 5, 0)

    def test_generate_values_rejects_wrong_direction(self):
        with self.assertRaises(ValueError):
            generate_sweep_values(1, 5, -1)


class BatchRunnerExecutionTest(unittest.TestCase):
    def _create_base_atp(self, atp_path: Path) -> None:
        content = (
            "BEGIN NEW DATA CASE\n"
            "/BRANCH\n"
            "  X0001AX0003A                5.   75.                                         0\n"
            "BLANK BRANCH\n"
            "BEGIN NEW DATA CASE\n"
            "BLANK\n"
        )
        atp_path.write_text(content, encoding="latin-1", errors="replace", newline="")

    def _read_branch_resistance(self, atp_path: Path) -> float:
        elements = parse_atp_file(atp_path)
        branch = next(element for element in elements if element.get("type") == "branch")
        return float(branch["resistance"])

    def _assert_expected_change(self, original_path: Path, generated_path: Path, expect_change: bool) -> None:
        original_lines = original_path.read_text(encoding="latin-1", errors="replace").splitlines()
        generated_lines = generated_path.read_text(encoding="latin-1", errors="replace").splitlines()

        self.assertEqual(len(original_lines), len(generated_lines))
        changed = [(idx, before, after) for idx, (before, after) in enumerate(zip(original_lines, generated_lines)) if before != after]

        if not expect_change:
            self.assertEqual(changed, [], "Quando o valor do sweep e igual ao original, o ATP deve permanecer identico")
            return

        self.assertEqual(len(changed), 1, "Cada ATP do sweep deve alterar somente uma linha")

        _line_index, before, after = changed[0]
        self.assertEqual(len(before), len(after), "A linha alterada nao pode mudar de tamanho")

        diff_positions = [idx for idx, (a, b) in enumerate(zip(before, after)) if a != b]
        self.assertTrue(diff_positions, "Era esperado ao menos um caractere alterado")
        self.assertLess(len(diff_positions), 10, "A alteracao deve ficar restrita ao campo numerico")

    def test_run_parameter_sweep_creates_isolated_runs(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            base_atp = root / "caso.atp"
            output_root = root / "out"
            self._create_base_atp(base_atp)

            parsed = parse_atp_file(base_atp)
            branch = next(element for element in parsed if element.get("type") == "branch")
            parameter = SweepParameterRef(
                line_index=int(branch["line_index"]),
                parameter="resistance",
                element_name="branch",
                label="Linha X0001A -> X0003A | Resistência",
            )

            parser_calls: list[tuple[int, float, str]] = []

            def fake_solver(atp_file_path: str, **_kwargs) -> str:
                atp_path = Path(atp_file_path)
                resistance = self._read_branch_resistance(atp_path)
                lis_path = atp_path.with_suffix(".lis")
                lis_path.write_text(
                    f"RESISTANCE={resistance}\n",
                    encoding="latin-1",
                    errors="replace",
                )
                return str(lis_path)

            def fake_lis_parser(
                lis_path: Path,
                run_dir: Path,
                value: float,
                run_index: int,
                total_runs: int,
            ) -> dict[str, object]:
                parser_calls.append((run_index, value, lis_path.read_text(encoding="latin-1").strip()))
                return {
                    "run_dir": str(run_dir),
                    "value": value,
                    "total_runs": total_runs,
                }

            summary = run_parameter_sweep(
                base_atp_path=base_atp,
                parameter_id=parameter,
                start=5,
                stop=15,
                step=5,
                output_dir=output_root,
                lis_parser=fake_lis_parser,
                solver_runner=fake_solver,
            )

            self.assertFalse(summary.cancelled)
            self.assertFalse(summary.stopped_on_error)
            self.assertEqual(summary.total_runs, 3)
            self.assertEqual(summary.success_count, 3)
            self.assertEqual(summary.failure_count, 0)
            self.assertEqual(len(parser_calls), 3)
            self.assertEqual(self._read_branch_resistance(base_atp), 5.0)

            expected_values = [5.0, 10.0, 15.0]
            self.assertEqual([result.value for result in summary.results], expected_values)

            for index, expected_value in enumerate(expected_values, start=1):
                result = summary.results[index - 1]
                self.assertEqual(result.status, "success")
                self.assertIsNotNone(result.atp_path)
                self.assertIsNotNone(result.lis_path)
                self.assertTrue(result.run_dir.exists())
                self.assertTrue(result.atp_path.exists())
                self.assertTrue(result.lis_path.exists())
                self.assertEqual(result.run_dir.name, f"run_{index:03d}_value_{int(expected_value)}")
                self.assertEqual(result.atp_path.name, "caso_param.atp")
                self.assertEqual(result.lis_path.name, "caso_param.lis")
                self.assertAlmostEqual(self._read_branch_resistance(result.atp_path), expected_value)
                self.assertIn(f"RESISTANCE={expected_value}", result.lis_path.read_text(encoding="latin-1"))
                self._assert_expected_change(base_atp, result.atp_path, expect_change=(expected_value != 5.0))

    def test_run_parameter_sweep_continues_after_failure_when_configured(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            base_atp = root / "caso.atp"
            output_root = root / "out"
            self._create_base_atp(base_atp)

            parsed = parse_atp_file(base_atp)
            branch = next(element for element in parsed if element.get("type") == "branch")

            def flaky_solver(atp_file_path: str, **_kwargs) -> str:
                atp_path = Path(atp_file_path)
                resistance = self._read_branch_resistance(atp_path)
                if resistance == 10.0:
                    raise RuntimeError("solver failure for 10")
                lis_path = atp_path.with_suffix(".lis")
                lis_path.write_text("OK\n", encoding="latin-1", errors="replace")
                return str(lis_path)

            summary = run_parameter_sweep(
                base_atp_path=base_atp,
                parameter_id={
                    "line_index": int(branch["line_index"]),
                    "parameter": "resistance",
                    "element_name": "branch",
                },
                start=5,
                stop=15,
                step=5,
                output_dir=output_root,
                solver_runner=flaky_solver,
                continue_on_error=True,
            )

            self.assertEqual(summary.success_count, 2)
            self.assertEqual(summary.failure_count, 1)
            self.assertEqual(summary.results[1].status, "failed")
            self.assertIn("solver failure for 10", summary.results[1].error or "")

    def test_run_parameter_sweep_parallelizes_isolated_runs(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            base_atp = root / "caso.atp"
            output_root = root / "out"
            include_dir = root / "includes"
            include_dir.mkdir(parents=True, exist_ok=True)
            (include_dir / "rede.inc").write_text("INCLUDE OK\n", encoding="latin-1", errors="replace")
            base_atp.write_text(
                (
                    "BEGIN NEW DATA CASE\n"
                    '$INSERT,"includes/rede.inc"\n'
                    "/BRANCH\n"
                    "  X0001AX0003A                5.   75.                                         0\n"
                    "BLANK BRANCH\n"
                ),
                encoding="latin-1",
                errors="replace",
                newline="",
            )

            parsed = parse_atp_file(base_atp)
            branch = next(element for element in parsed if element.get("type") == "branch")
            parameter = SweepParameterRef(
                line_index=int(branch["line_index"]),
                parameter="resistance",
                element_name="branch",
            )

            def slow_solver(atp_file_path: str, **_kwargs) -> str:
                atp_path = Path(atp_file_path)
                include_copy = atp_path.parent / "includes" / "rede.inc"
                if not include_copy.exists():
                    raise FileNotFoundError(f"Include nao copiado para workspace isolado: {include_copy}")

                resistance = self._read_branch_resistance(atp_path)
                time.sleep(0.25)
                lis_path = atp_path.with_suffix(".lis")
                lis_path.write_text(
                    f"RESISTANCE={resistance}\n",
                    encoding="latin-1",
                    errors="replace",
                )
                return str(lis_path)

            started = time.monotonic()
            summary = run_parameter_sweep(
                base_atp_path=base_atp,
                parameter_id=parameter,
                start=5,
                stop=20,
                step=5,
                output_dir=output_root,
                solver_runner=slow_solver,
                max_parallel_runs=4,
            )
            elapsed = time.monotonic() - started

            self.assertLess(elapsed, 0.9, "Sweep paralelo deveria reduzir o tempo total")
            self.assertEqual(summary.total_runs, 4)
            self.assertEqual(summary.success_count, 4)
            self.assertEqual(summary.failure_count, 0)

            for expected_value, result in zip([5.0, 10.0, 15.0, 20.0], summary.results):
                self.assertEqual(result.status, "success")
                self.assertIsNotNone(result.atp_path)
                self.assertIsNotNone(result.lis_path)
                self.assertTrue(result.atp_path.exists())
                self.assertTrue(result.lis_path.exists())
                self.assertAlmostEqual(self._read_branch_resistance(result.atp_path), expected_value)


    def test_cancel_marks_unstarted_runs_and_reports_partial_progress(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            base_atp = root / "caso.atp"
            self._create_base_atp(base_atp)
            branch = parse_atp_file(base_atp)[0]
            cancel_event = threading.Event()
            events: list[dict] = []

            def fake_solver(atp_file_path: str, **_kwargs) -> str:
                lis_path = Path(atp_file_path).with_suffix(".lis")
                lis_path.write_text("OK\n", encoding="latin-1")
                return str(lis_path)

            def on_event(event: dict) -> None:
                events.append(event)
                if event.get("type") == "run_succeeded":
                    cancel_event.set()

            summary = run_parameter_sweep(
                base_atp_path=base_atp,
                parameter_id={"line_index": branch["line_index"], "parameter": "resistance"},
                start=5,
                stop=15,
                step=5,
                output_dir=root / "out",
                solver_runner=fake_solver,
                cancel_event=cancel_event,
                event_callback=on_event,
            )

            self.assertTrue(summary.cancelled)
            self.assertEqual(
                [result.status for result in summary.results],
                ["success", "cancelled", "cancelled"],
            )
            self.assertEqual(summary.processed_count, 1)
            self.assertEqual(summary.cancelled_count, 2)
            self.assertAlmostEqual(events[-1]["progress"], 1 / 3)
            self.assertTrue(summary.output_dir.exists())
            self.assertTrue(summary.results[0].run_dir.exists())
            self.assertFalse(summary.results[1].run_dir.exists())
            self.assertFalse(summary.results[2].run_dir.exists())
            self.assertIsNone(summary.results[1].atp_path)
            self.assertIsNone(summary.results[1].lis_path)
            self.assertIsNone(summary.results[2].atp_path)
            self.assertIsNone(summary.results[2].lis_path)

    def test_cancel_reaches_solver_already_in_progress(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            base_atp = root / "caso.atp"
            self._create_base_atp(base_atp)
            branch = parse_atp_file(base_atp)[0]
            cancel_event = threading.Event()
            solver_started = threading.Event()

            def cancellable_solver(atp_file_path: str, **kwargs) -> str:
                solver_started.set()
                received_cancel_event = kwargs.get("cancel_event")
                while not received_cancel_event.is_set():
                    time.sleep(0.01)
                raise RuntimeError("cancelled active solver")

            def request_cancel() -> None:
                solver_started.wait(timeout=1)
                cancel_event.set()

            threading.Thread(target=request_cancel, daemon=True).start()
            started = time.monotonic()
            summary = run_parameter_sweep(
                base_atp_path=base_atp,
                parameter_id={"line_index": branch["line_index"], "parameter": "resistance"},
                start=5,
                stop=15,
                step=5,
                output_dir=root / "out",
                solver_runner=cancellable_solver,
                cancel_event=cancel_event,
            )

            self.assertLess(time.monotonic() - started, 1.0)
            self.assertTrue(summary.cancelled)
            self.assertEqual(
                [result.status for result in summary.results],
                ["cancelled", "cancelled", "cancelled"],
            )
            self.assertFalse(summary.output_dir.exists())
            self.assertTrue(all(not result.run_dir.exists() for result in summary.results))

    def test_parallel_solver_serializes_lis_postprocessing(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            base_atp = root / "caso.atp"
            self._create_base_atp(base_atp)
            branch = parse_atp_file(base_atp)[0]
            state_lock = threading.Lock()
            active_parsers = 0
            maximum_active_parsers = 0

            def fake_solver(atp_file_path: str, **_kwargs) -> str:
                lis_path = Path(atp_file_path).with_suffix(".lis")
                lis_path.write_text("OK\n", encoding="latin-1")
                return str(lis_path)

            def guarded_parser(*_args):
                nonlocal active_parsers, maximum_active_parsers
                with state_lock:
                    active_parsers += 1
                    maximum_active_parsers = max(maximum_active_parsers, active_parsers)
                time.sleep(0.03)
                with state_lock:
                    active_parsers -= 1

            summary = run_parameter_sweep(
                base_atp_path=base_atp,
                parameter_id={"line_index": branch["line_index"], "parameter": "resistance"},
                start=5,
                stop=20,
                step=5,
                output_dir=root / "out",
                solver_runner=fake_solver,
                lis_parser=guarded_parser,
                max_parallel_runs=4,
            )

            self.assertEqual(summary.success_count, 4)
            self.assertEqual(maximum_active_parsers, 1)


if __name__ == "__main__":
    unittest.main()
