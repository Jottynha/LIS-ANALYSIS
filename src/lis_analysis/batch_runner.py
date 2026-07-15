from __future__ import annotations

import concurrent.futures
import shutil
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Mapping

from .atp_parser import parse_atp_file_cached
from .atp_writer import apply_parameter_overrides
from .solver.atp_runner import (
    cleanup_staged_atp_result,
    iter_staged_atp_artifacts,
    run_atp_solver,
)

SweepEventCallback = Callable[[dict[str, Any]], None]
SweepLisParser = Callable[[Path, Path, float, int, int], Any]


@dataclass(frozen=True)
class SweepParameterRef:
    line_index: int
    parameter: str
    element_name: str = "element"
    label: str = ""

    @property
    def display_label(self) -> str:
        if self.label:
            return self.label
        return f"Linha {self.line_index + 1} | {self.parameter}"


@dataclass
class SweepRunResult:
    run_index: int
    value: float
    run_dir: Path
    atp_path: Path | None = None
    lis_path: Path | None = None
    status: str = "pending"
    elapsed_seconds: float = 0.0
    error: str | None = None
    analysis: Any = None


@dataclass
class SweepExecutionSummary:
    parameter: SweepParameterRef
    values: list[float]
    output_dir: Path
    started_at: datetime
    finished_at: datetime | None = None
    cancelled: bool = False
    stopped_on_error: bool = False
    results: list[SweepRunResult] = field(default_factory=list)

    @property
    def total_runs(self) -> int:
        return len(self.values)

    @property
    def success_count(self) -> int:
        return sum(1 for result in self.results if result.status == "success")

    @property
    def failure_count(self) -> int:
        return sum(1 for result in self.results if result.status == "failed")

    @property
    def cancelled_count(self) -> int:
        return sum(1 for result in self.results if result.status == "cancelled")

    @property
    def skipped_count(self) -> int:
        return sum(1 for result in self.results if result.status == "skipped")

    @property
    def processed_count(self) -> int:
        """Quantidade de runs que realmente chegaram a uma conclusao de execucao."""
        return self.success_count + self.failure_count

    @property
    def elapsed_seconds(self) -> float:
        if self.finished_at is None:
            return 0.0
        return max(0.0, (self.finished_at - self.started_at).total_seconds())


@dataclass(frozen=True)
class _PreparedSweepParameter:
    line_index: int
    parameter: str
    start: int
    end: int
    editable: bool
    original_value: float


def generate_sweep_values(start: float, stop: float, step: float) -> list[float]:
    """Gera valores do sweep incluindo o fim quando ele cair exatamente na sequencia."""
    start_dec = _to_decimal(start)
    stop_dec = _to_decimal(stop)
    step_dec = _to_decimal(step)

    if step_dec == 0:
        raise ValueError("step nao pode ser zero")

    if start_dec < stop_dec and step_dec <= 0:
        raise ValueError("step deve ser positivo quando start < stop")

    if start_dec > stop_dec and step_dec >= 0:
        raise ValueError("step deve ser negativo quando start > stop")

    if start_dec == stop_dec:
        return [float(start_dec)]

    values: list[float] = []
    current = start_dec
    max_iterations = 100000

    for _ in range(max_iterations):
        if step_dec > 0 and current > stop_dec:
            break
        if step_dec < 0 and current < stop_dec:
            break

        values.append(float(current))
        current += step_dec
    else:
        raise RuntimeError("Sweep excedeu o limite de iteracoes; verifique start/stop/step")

    return values


def run_parameter_sweep(
    base_atp_path: str | Path,
    parameter_id: SweepParameterRef | Mapping[str, Any] | str,
    start: float,
    stop: float,
    step: float,
    output_dir: str | Path,
    *,
    lis_parser: SweepLisParser | None = None,
    solver_runner: Callable[..., str] = run_atp_solver,
    solver_timeout: int = 600,
    continue_on_error: bool = True,
    cancel_event: Any | None = None,
    event_callback: SweepEventCallback | None = None,
    max_parallel_runs: int = 1,
) -> SweepExecutionSummary:
    """Executa sweep parametrico com opcao de isolamento/paralelismo por run."""
    base_path = Path(base_atp_path)
    if not base_path.exists():
        raise FileNotFoundError(f"Arquivo .atp nao encontrado: {base_path}")

    parameter_ref = _normalize_parameter_ref(parameter_id)
    values = generate_sweep_values(start, stop, step)
    if not values:
        raise ValueError("Nenhum valor foi gerado para o sweep")

    sweep_root = Path(output_dir)
    sweep_root.mkdir(parents=True, exist_ok=True)
    sweep_dir = sweep_root / f"sweep_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    sweep_dir.mkdir(parents=True, exist_ok=True)

    summary = SweepExecutionSummary(
        parameter=parameter_ref,
        values=values,
        output_dir=sweep_dir,
        started_at=datetime.now(),
    )
    total_runs = len(values)
    source_elements = parse_atp_file_cached(base_path)
    original_lines = _read_text_lines_preserve_newlines(base_path)
    prepared_parameter = _prepare_sweep_parameter(source_elements, parameter_ref)

    summary.results = [
        SweepRunResult(
            run_index=run_index,
            value=float(value),
            run_dir=sweep_dir / f"run_{run_index:03d}_value_{_format_value_for_path(value)}",
        )
        for run_index, value in enumerate(values, start=1)
    ]

    parallel_runs = _resolve_parallel_runs(max_parallel_runs, total_runs)
    if parallel_runs > 1 and not continue_on_error:
        parallel_runs = 1
        _emit_event(
            event_callback,
            type="sweep_mode_adjusted",
            message=(
                "Execucao paralela desabilitada porque 'parar ao primeiro erro' "
                "exige ordem estritamente sequencial."
            ),
            run_index=0,
            total_runs=total_runs,
            value=None,
            progress=0.0,
            output_dir=sweep_dir,
        )

    isolated_dependency_files: tuple[tuple[Path, Path], ...] = ()
    if parallel_runs > 1:
        dependency_plan = _collect_isolated_workspace_dependencies(base_path)
        if dependency_plan is None:
            parallel_runs = 1
            _emit_event(
                event_callback,
                type="sweep_mode_adjusted",
                message=(
                    "Execucao paralela desabilitada: foram detectados $INSERTs "
                    "relativos fora da pasta base do ATP."
                ),
                run_index=0,
                total_runs=total_runs,
                value=None,
                progress=0.0,
                output_dir=sweep_dir,
            )
        else:
            isolated_dependency_files = dependency_plan

    _emit_event(
        event_callback,
        type="sweep_started",
        message=(
            f"Iniciando sweep de {parameter_ref.display_label} "
            f"com {total_runs} execucao(oes)"
        ),
        run_index=0,
        total_runs=total_runs,
        value=None,
        progress=0.0,
        output_dir=sweep_dir,
    )
    if parallel_runs > 1:
        _emit_event(
            event_callback,
            type="sweep_parallel_enabled",
            message=f"Execucao paralela habilitada com {parallel_runs} worker(s).",
            run_index=0,
            total_runs=total_runs,
            value=None,
            progress=0.0,
            output_dir=sweep_dir,
        )

    def _run_and_finalize(result: SweepRunResult, *, isolate_workspace: bool) -> None:
        _execute_sweep_run(
            base_path=base_path,
            prepared_parameter=prepared_parameter,
            original_lines=original_lines,
            result=result,
            total_runs=total_runs,
            solver_runner=solver_runner,
            solver_timeout=solver_timeout,
            event_callback=event_callback,
            cancel_event=cancel_event,
            isolate_workspace=isolate_workspace,
            isolated_dependency_files=isolated_dependency_files,
        )
        _finalize_sweep_run(
            result=result,
            total_runs=total_runs,
            lis_parser=lis_parser,
            event_callback=event_callback,
            lis_parser_lock=lis_parser_lock,
            cancel_event=cancel_event,
        )

    # O solver pode rodar em paralelo, mas o pos-processamento da GUI usa
    # matplotlib.pyplot, cujo estado global nao e seguro entre threads.
    lis_parser_lock = threading.Lock() if parallel_runs > 1 and lis_parser is not None else None

    if parallel_runs <= 1:
        for result in summary.results:
            if cancel_event is not None and cancel_event.is_set():
                summary.cancelled = True
                break

            _run_and_finalize(result, isolate_workspace=False)
            if result.status == "failed" and not continue_on_error:
                summary.stopped_on_error = True
                break
    else:
        with concurrent.futures.ThreadPoolExecutor(max_workers=parallel_runs) as executor:
            future_map = {
                executor.submit(_run_and_finalize, result, isolate_workspace=True): result
                for result in summary.results
            }
            for future in concurrent.futures.as_completed(future_map):
                future.result()

    if cancel_event is not None and cancel_event.is_set():
        summary.cancelled = True

    if summary.cancelled:
        for result in summary.results:
            if result.status == "pending":
                result.status = "cancelled"
                result.error = "cancelled before execution"
    elif summary.stopped_on_error:
        for result in summary.results:
            if result.status == "pending":
                result.status = "skipped"
                result.error = "not executed after previous failure"

    summary.finished_at = datetime.now()
    processed_progress = summary.processed_count / max(1, total_runs)

    if summary.cancelled:
        for cancelled_result in summary.results:
            if cancelled_result.status != "cancelled":
                continue
            shutil.rmtree(cancelled_result.run_dir, ignore_errors=True)
            cancelled_result.atp_path = None
            cancelled_result.lis_path = None

        try:
            sweep_dir.rmdir()
        except OSError:
            pass

        _emit_event(
            event_callback,
            type="sweep_cancelled",
            message="Sweep cancelado pelo usuario",
            run_index=summary.processed_count,
            total_runs=total_runs,
            value=None,
            progress=processed_progress,
            output_dir=sweep_dir,
        )

    _emit_event(
        event_callback,
        type="sweep_finished",
        message=(
            f"Sweep finalizado: {summary.success_count} sucesso(s), "
            f"{summary.failure_count} falha(s), {summary.cancelled_count} cancelada(s), "
            f"{summary.skipped_count} ignorada(s)"
        ),
        run_index=summary.processed_count,
        total_runs=total_runs,
        value=None,
        progress=1.0 if not summary.cancelled and not summary.stopped_on_error else processed_progress,
        output_dir=sweep_dir,
        summary=summary,
    )

    return summary


def _to_decimal(value: float | str | Decimal) -> Decimal:
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"Valor numerico invalido: {value}") from exc


def _normalize_parameter_ref(
    parameter_id: SweepParameterRef | Mapping[str, Any] | str,
) -> SweepParameterRef:
    if isinstance(parameter_id, SweepParameterRef):
        return parameter_id

    if isinstance(parameter_id, str):
        line_text, sep, parameter = parameter_id.partition(":")
        if not sep:
            raise ValueError("parameter_id em texto deve seguir o formato 'line_index:parametro'")
        return SweepParameterRef(line_index=int(line_text), parameter=parameter.strip())

    if isinstance(parameter_id, Mapping):
        if "line_index" not in parameter_id or "parameter" not in parameter_id:
            raise ValueError("parameter_id deve conter 'line_index' e 'parameter'")
        return SweepParameterRef(
            line_index=int(parameter_id["line_index"]),
            parameter=str(parameter_id["parameter"]),
            element_name=str(parameter_id.get("element_name", parameter_id.get("name", "element"))),
            label=str(parameter_id.get("label", "")),
        )

    raise TypeError("parameter_id invalido")


def _resolve_parallel_runs(requested_parallel_runs: int, total_runs: int) -> int:
    try:
        requested = int(requested_parallel_runs)
    except (TypeError, ValueError) as exc:
        raise ValueError("max_parallel_runs deve ser um inteiro >= 1") from exc

    if requested < 1:
        raise ValueError("max_parallel_runs deve ser >= 1")

    return min(requested, max(1, total_runs))


def _prepare_sweep_parameter(
    elements: list[dict[str, Any]],
    parameter_ref: SweepParameterRef,
) -> _PreparedSweepParameter:
    for element in elements:
        if int(element.get("line_index", -1)) != parameter_ref.line_index:
            continue

        params = element.get("parameters")
        if not isinstance(params, dict) or parameter_ref.parameter not in params:
            break

        meta = params[parameter_ref.parameter]
        if not isinstance(meta, dict):
            break

        return _PreparedSweepParameter(
            line_index=parameter_ref.line_index,
            parameter=parameter_ref.parameter,
            start=int(meta.get("start", -1)),
            end=int(meta.get("end", -1)),
            editable=bool(meta.get("editable", True)),
            original_value=float(meta.get("original_value", meta.get("value"))),
        )

    raise ValueError(
        f"Parametro '{parameter_ref.parameter}' nao encontrado na linha {parameter_ref.line_index + 1}"
    )


def _render_sweep_atp_lines(
    original_lines: list[str],
    prepared_parameter: _PreparedSweepParameter,
    value: float,
) -> list[str]:
    overrides = {
        prepared_parameter.line_index: [
            {
                "field": prepared_parameter.parameter,
                "start": prepared_parameter.start,
                "end": prepared_parameter.end,
                "old_value": prepared_parameter.original_value,
                "new_value": float(value),
                "editable": prepared_parameter.editable,
            }
        ]
    }
    return apply_parameter_overrides(original_lines, overrides)


def _write_rendered_atp_file(output_path: Path, rendered_lines: list[str]) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="latin-1", errors="replace", newline="") as file:
        file.write("".join(rendered_lines))
    return output_path


def _parse_insert_targets_from_file(path: Path) -> list[str]:
    try:
        lines = path.read_text(errors="replace").splitlines()
    except Exception:
        return []

    targets: list[str] = []
    for raw_line in lines:
        stripped = raw_line.strip()
        if not stripped.upper().startswith("$INSERT"):
            continue

        parts = stripped.split(",", 1)
        if len(parts) < 2:
            continue

        target = parts[1].strip().strip('"').strip("'")
        if target:
            targets.append(target)
    return targets


def _is_windows_absolute_path(target: str) -> bool:
    if len(target) < 3:
        return False
    return target[1] == ":" and target[2] in ("\\", "/") and target[0].isalpha()


def _collect_isolated_workspace_dependencies(
    base_atp_path: Path,
) -> tuple[tuple[Path, Path], ...] | None:
    base_root = base_atp_path.parent.resolve()
    queue = [base_atp_path.resolve()]
    visited_files: set[Path] = set()
    planned_copies: dict[Path, Path] = {}

    while queue:
        current = queue.pop()
        if current in visited_files:
            continue
        visited_files.add(current)

        for target in _parse_insert_targets_from_file(current):
            if _is_windows_absolute_path(target):
                continue

            normalized = Path(target.replace("\\", "/"))
            if normalized.is_absolute():
                continue

            resolved = (current.parent / normalized).resolve()
            try:
                relative_to_base = resolved.relative_to(base_root)
            except ValueError:
                return None

            if resolved not in planned_copies:
                planned_copies[resolved] = relative_to_base
                if resolved.exists():
                    queue.append(resolved)

    ordered = sorted(planned_copies.items(), key=lambda item: str(item[1]).lower())
    return tuple((source, relative_path) for source, relative_path in ordered)


def _populate_isolated_workspace(
    workspace_dir: Path,
    dependency_files: tuple[tuple[Path, Path], ...],
) -> None:
    workspace_dir.mkdir(parents=True, exist_ok=True)
    for source, relative_path in dependency_files:
        if not source.exists():
            continue
        target = workspace_dir / relative_path
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(source), str(target))


def _execute_sweep_run(
    *,
    base_path: Path,
    prepared_parameter: _PreparedSweepParameter,
    original_lines: list[str],
    result: SweepRunResult,
    total_runs: int,
    solver_runner: Callable[..., str],
    solver_timeout: int,
    event_callback: SweepEventCallback | None,
    cancel_event: Any | None,
    isolate_workspace: bool,
    isolated_dependency_files: tuple[tuple[Path, Path], ...],
) -> None:
    result.run_dir.mkdir(parents=True, exist_ok=True)

    if cancel_event is not None and cancel_event.is_set():
        result.status = "cancelled"
        result.error = "cancelled"
        return

    execution_atp_path = _build_execution_atp_path(
        base_path if not isolate_workspace else (result.run_dir / "_solver_workspace" / base_path.name),
        result.run_index,
    )
    workspace_dir = execution_atp_path.parent if isolate_workspace else None
    generated_lis_path: Path | None = None
    run_started = time.monotonic()

    _emit_event(
        event_callback,
        type="run_started",
        message=f"Run {result.run_index}/{total_runs}: preparando valor {result.value:g}",
        run_index=result.run_index,
        total_runs=total_runs,
        value=result.value,
        progress=(result.run_index - 1) / total_runs,
        run_dir=result.run_dir,
    )

    try:
        if workspace_dir is not None:
            _populate_isolated_workspace(workspace_dir, isolated_dependency_files)

        rendered_lines = _render_sweep_atp_lines(original_lines, prepared_parameter, result.value)
        _write_rendered_atp_file(execution_atp_path, rendered_lines)

        _emit_event(
            event_callback,
            type="atp_written",
            message=f"ATP parametrizado salvo: {execution_atp_path.name}",
            run_index=result.run_index,
            total_runs=total_runs,
            value=result.value,
            progress=((result.run_index - 1) + 0.15) / total_runs,
            run_dir=result.run_dir,
            atp_path=execution_atp_path,
        )

        def _solver_status_callback(message: str) -> None:
            _emit_event(
                event_callback,
                type="solver_message",
                message=message,
                run_index=result.run_index,
                total_runs=total_runs,
                value=result.value,
                progress=((result.run_index - 1) + 0.55) / total_runs,
                run_dir=result.run_dir,
            )

        def _solver_progress_callback(progress: float, detail: str) -> None:
            normalized = max(0.0, min(1.0, float(progress)))
            _emit_event(
                event_callback,
                type="solver_progress",
                message=detail,
                run_index=result.run_index,
                total_runs=total_runs,
                value=result.value,
                simulation_progress=normalized,
                progress=((result.run_index - 1) + normalized) / total_runs,
                run_dir=result.run_dir,
            )

        generated_lis_path = Path(
            solver_runner(
                str(execution_atp_path),
                timeout=solver_timeout,
                status_callback=_solver_status_callback,
                cancel_event=cancel_event,
                progress_callback=_solver_progress_callback,
            )
        )

        if cancel_event is not None and cancel_event.is_set():
            raise RuntimeError("cancelled by user")

        _emit_event(
            event_callback,
            type="solver_finished",
            message=f"Solver finalizado para valor {result.value:g}",
            run_index=result.run_index,
            total_runs=total_runs,
            value=result.value,
            progress=((result.run_index - 1) + 0.8) / total_runs,
            run_dir=result.run_dir,
            lis_path=generated_lis_path,
        )

        relocated = _relocate_run_artifacts(
            base_atp_path=base_path,
            execution_atp_path=execution_atp_path,
            generated_lis_path=generated_lis_path,
            run_dir=result.run_dir,
        )
        result.atp_path = relocated["atp_path"]
        result.lis_path = relocated["lis_path"]
        result.elapsed_seconds = time.monotonic() - run_started
        result.status = "solver_completed"

        _emit_event(
            event_callback,
            type="lis_ready",
            message=f"LIS pronto em {result.lis_path}",
            run_index=result.run_index,
            total_runs=total_runs,
            value=result.value,
            progress=((result.run_index - 1) + 0.9) / total_runs,
            run_dir=result.run_dir,
            lis_path=result.lis_path,
        )
    except Exception as exc:
        result.elapsed_seconds = time.monotonic() - run_started
        result.error = str(exc)
        result.status = "cancelled" if cancel_event is not None and cancel_event.is_set() else "failed"

        relocated = _relocate_run_artifacts(
            base_atp_path=base_path,
            execution_atp_path=execution_atp_path,
            generated_lis_path=generated_lis_path,
            run_dir=result.run_dir,
        )
        result.atp_path = relocated.get("atp_path")
        result.lis_path = relocated.get("lis_path")

        was_cancelled = cancel_event is not None and cancel_event.is_set()
        _emit_event(
            event_callback,
            type="run_cancelled" if was_cancelled else "run_failed",
            message=(
                f"Run {result.run_index}/{total_runs} "
                f"{'cancelado' if was_cancelled else 'falhou'} em "
                f"{result.elapsed_seconds:.2f}s: {result.error}"
            ),
            run_index=result.run_index,
            total_runs=total_runs,
            value=result.value,
            progress=result.run_index / total_runs,
            run_dir=result.run_dir,
            error=result.error,
        )
    finally:
        cleanup_staged_atp_result(generated_lis_path)
        if workspace_dir is not None:
            shutil.rmtree(workspace_dir, ignore_errors=True)


def _finalize_sweep_run(
    *,
    result: SweepRunResult,
    total_runs: int,
    lis_parser: SweepLisParser | None,
    event_callback: SweepEventCallback | None,
    lis_parser_lock: threading.Lock | None,
    cancel_event: Any | None,
) -> None:
    if result.status != "solver_completed":
        return
    if cancel_event is not None and cancel_event.is_set():
        result.status = "cancelled"
        result.error = "cancelled before post-processing"
        return

    try:
        if lis_parser is not None and result.lis_path is not None:
            _emit_event(
                event_callback,
                type="lis_parsing",
                message=f"Pos-processando {result.lis_path.name}",
                run_index=result.run_index,
                total_runs=total_runs,
                value=result.value,
                progress=((result.run_index - 1) + 0.95) / total_runs,
                run_dir=result.run_dir,
                lis_path=result.lis_path,
            )
            parse_started = time.monotonic()
            def _parse_lis() -> Any:
                return lis_parser(
                    result.lis_path,
                    result.run_dir,
                    result.value,
                    result.run_index,
                    total_runs,
                )

            if lis_parser_lock is None:
                result.analysis = _parse_lis()
            else:
                with lis_parser_lock:
                    result.analysis = _parse_lis()
            result.elapsed_seconds += time.monotonic() - parse_started

        if cancel_event is not None and cancel_event.is_set():
            result.status = "cancelled"
            result.error = "cancelled during post-processing"
            return

        result.status = "success"
        _emit_event(
            event_callback,
            type="run_succeeded",
            message=(
                f"Run {result.run_index}/{total_runs} concluido em "
                f"{result.elapsed_seconds:.2f}s"
            ),
            run_index=result.run_index,
            total_runs=total_runs,
            value=result.value,
            progress=result.run_index / total_runs,
            run_dir=result.run_dir,
            lis_path=result.lis_path,
        )
    except Exception as exc:
        result.error = str(exc)
        result.status = "failed"
        _emit_event(
            event_callback,
            type="run_failed",
            message=(
                f"Run {result.run_index}/{total_runs} falhou em "
                f"{result.elapsed_seconds:.2f}s: {result.error}"
            ),
            run_index=result.run_index,
            total_runs=total_runs,
            value=result.value,
            progress=result.run_index / total_runs,
            run_dir=result.run_dir,
            error=result.error,
        )


def _emit_event(callback: SweepEventCallback | None, **payload: Any) -> None:
    if callback is None:
        return
    try:
        callback(payload)
    except Exception:
        return


def _build_execution_atp_path(base_atp_path: Path, run_index: int) -> Path:
    return base_atp_path.parent / f"{base_atp_path.stem}__sweep_{run_index:03d}{base_atp_path.suffix}"


def _read_text_lines_preserve_newlines(path: Path) -> list[str]:
    with path.open("r", encoding="latin-1", errors="replace", newline="") as file:
        return file.read().splitlines(keepends=True)


def _format_value_for_path(value: float) -> str:
    raw = f"{float(value):.10g}"
    return (
        raw.replace("-", "neg_")
        .replace("+", "")
        .replace(".", "p")
    )


def _unique_destination(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _move_or_copy(src: Path, dest: Path) -> Path:
    if not src.exists():
        return dest

    final_dest = _unique_destination(dest)
    final_dest.parent.mkdir(parents=True, exist_ok=True)

    try:
        return Path(shutil.move(str(src), str(final_dest)))
    except Exception:
        shutil.copy2(str(src), str(final_dest))
        return final_dest


def _relocate_run_artifacts(
    *,
    base_atp_path: Path,
    execution_atp_path: Path,
    generated_lis_path: Path | None,
    run_dir: Path,
) -> dict[str, Path | None]:
    run_dir.mkdir(parents=True, exist_ok=True)

    param_stem = f"{base_atp_path.stem}_param"
    atp_target = run_dir / f"{param_stem}{base_atp_path.suffix}"
    lis_target = run_dir / f"{param_stem}{generated_lis_path.suffix}" if generated_lis_path else None

    relocated_atp: Path | None = None
    relocated_lis: Path | None = None
    moved_sources: set[Path] = set()

    if execution_atp_path.exists():
        relocated_atp = _move_or_copy(execution_atp_path, atp_target)
        moved_sources.add(execution_atp_path.resolve())

    staged_artifacts = (
        iter_staged_atp_artifacts(generated_lis_path)
        if generated_lis_path is not None
        else ()
    )
    if generated_lis_path is not None and generated_lis_path.exists() and lis_target is not None:
        relocated_lis = _move_or_copy(generated_lis_path, lis_target)
        moved_sources.add(generated_lis_path.resolve())

    for sidecar in staged_artifacts:
        if generated_lis_path is not None and sidecar == generated_lis_path:
            continue
        if not sidecar.exists():
            continue
        sidecar_target = run_dir / f"{param_stem}{sidecar.suffix}"
        _move_or_copy(sidecar, sidecar_target)

    cleanup_staged_atp_result(generated_lis_path)

    for sidecar in execution_atp_path.parent.glob(f"{execution_atp_path.stem}.*"):
        try:
            resolved = sidecar.resolve()
        except Exception:
            continue
        if resolved in moved_sources or not sidecar.is_file():
            continue
        sidecar_target = run_dir / f"{param_stem}{sidecar.suffix}"
        _move_or_copy(sidecar, sidecar_target)
        moved_sources.add(resolved)

    return {
        "atp_path": relocated_atp,
        "lis_path": relocated_lis,
    }
