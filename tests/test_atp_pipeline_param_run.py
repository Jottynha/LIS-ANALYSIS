from __future__ import annotations

import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from atp_parser import parse_atp_file, update_parameter
from atp_writer import write_atp_file
from solver.atp_runner import run_atp_solver


class _FakeStdin:
    def write(self, _data: str) -> int:
        return 0

    def flush(self) -> None:
        return

    def close(self) -> None:
        return


class _FakeStdout:
    def __init__(self, lines: list[str]):
        self._lines = lines
        self._index = 0

    def __iter__(self):
        return self

    def __next__(self) -> str:
        if self._index >= len(self._lines):
            raise StopIteration
        line = self._lines[self._index]
        self._index += 1
        return line


def _make_fake_popen_class(
    lis_name: str,
    lis_content: str,
    finish_after_sec: float = 1.0,
):
    class FakePopen:
        def __init__(self, command, cwd=None, **_kwargs):
            self.command = command
            self.cwd = Path(cwd)
            self.pid = 99999
            self.stdin = _FakeStdin()
            self.stdout = _FakeStdout(["Total execution time was 0.1 seconds\n"])
            self._start = time.time()
            self._terminated = False
            self._killed = False

            def _emit_lis() -> None:
                time.sleep(0.2)
                target = self.cwd / lis_name
                target.write_text(lis_content, encoding="latin-1", errors="replace")

            threading.Thread(target=_emit_lis, daemon=True).start()
            self._finish_after_sec = finish_after_sec

        def poll(self):
            if self._terminated:
                return -9 if self._killed else 0
            if time.time() - self._start >= self._finish_after_sec:
                return 0
            return None

        def wait(self, timeout=None):
            start = time.time()
            while True:
                rc = self.poll()
                if rc is not None:
                    return rc
                if timeout is not None and (time.time() - start) > timeout:
                    raise TimeoutError("fake process wait timeout")
                time.sleep(0.02)

        def terminate(self):
            self._terminated = True

        def kill(self):
            self._terminated = True
            self._killed = True

    return FakePopen


class ATPPipelineParamRunTest(unittest.TestCase):
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

    def _create_parametrized_atp(self, original_atp: Path, output_atp: Path) -> None:
        elements = parse_atp_file(original_atp)
        branch = next(e for e in elements if e.get("type") == "branch")
        line_index = int(branch["line_index"])

        update_parameter(
            elements,
            element_name="branch",
            new_value=7.0,
            line_index=line_index,
            parameter_name="resistance",
        )

        with original_atp.open("r", encoding="latin-1", errors="replace", newline="") as f:
            original_lines = f.read().splitlines(keepends=True)

        write_atp_file(elements, original_lines, output_atp)

    def test_param_run_detects_recent_lis_with_alternative_name(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            original = tmp / "caso.atp"
            parametrized = tmp / "caso__param.atp"
            self._create_base_atp(original)
            self._create_parametrized_atp(original, parametrized)

            fake_popen = _make_fake_popen_class(
                lis_name="resultado_alternativo.lis",
                lis_content="LIS OK\n",
            )

            with patch("solver.atp_runner.subprocess.Popen", new=fake_popen):
                started = time.monotonic()
                lis_path = run_atp_solver(str(parametrized), timeout=30)
                elapsed = time.monotonic() - started

            self.assertLess(elapsed, 20.0, "Execução demorou além do esperado (possível loop)")
            self.assertTrue(Path(lis_path).exists())
            self.assertEqual(Path(lis_path).name, "resultado_alternativo.lis")

    def test_param_run_raises_clear_error_on_kill_lis(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            original = tmp / "caso.atp"
            parametrized = tmp / "caso__param.atp"
            self._create_base_atp(original)
            self._create_parametrized_atp(original, parametrized)

            kill_lis = (
                "--------------------------------------------------------------------------------\n"
                " <<<< EMTP error stop.  KILL code #      Overlay number      Nearby statement\n"
                "               273                 1                     1832\n"
                "KILL = 273.  SUBROUTINE UNIX ...\n"
                "--------------------------------------------------------------------------------\n"
            )
            fake_popen = _make_fake_popen_class(
                lis_name="resultado_kill.lis",
                lis_content=kill_lis,
            )

            with patch("solver.atp_runner.subprocess.Popen", new=fake_popen):
                with self.assertRaises(RuntimeError) as ctx:
                    run_atp_solver(str(parametrized), timeout=30)

            message = str(ctx.exception)
            self.assertIn("KILL", message)
            self.assertIn("Trecho do LIS", message)


if __name__ == "__main__":
    unittest.main()
