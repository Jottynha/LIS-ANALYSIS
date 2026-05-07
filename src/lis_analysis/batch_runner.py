from __future__ import annotations

import shutil
import time
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Mapping

from .atp_parser import parse_atp_file_cached, update_parameter
from .atp_writer import write_atp_file
from .solver.atp_runner import run_atp_solver

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
    def elapsed_seconds(self) -> float:
        if self.finished_at is None:
            return 0.0
        return max(0.0, (self.finished_at - self.started_at).total_seconds())


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
) -> SweepExecutionSummary:
    """Executa um sweep parametrico sequencial, isolando cada run em sua propria pasta."""
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

    for run_index, value in enumerate(values, start=1):
        if cancel_event is not None and cancel_event.is_set():
            summary.cancelled = True
            break

        run_dir = sweep_dir / f"run_{run_index:03d}_value_{_format_value_for_path(value)}"
        run_dir.mkdir(parents=True, exist_ok=True)

        result = SweepRunResult(run_index=run_index, value=float(value), run_dir=run_dir)
        summary.results.append(result)

        execution_atp_path = _build_execution_atp_path(base_path, run_index)
        generated_lis_path: Path | None = None
        run_started = time.monotonic()

        _emit_event(
            event_callback,
            type="run_started",
            message=f"Run {run_index}/{total_runs}: preparando valor {value:g}",
            run_index=run_index,
            total_runs=total_runs,
            value=float(value),
            progress=(run_index - 1) / total_runs,
            run_dir=run_dir,
        )

        try:
            elements = parse_atp_file_cached(base_path)
            original_lines = _read_text_lines_preserve_newlines(base_path)

            update_parameter(
                elements,
                element_name=parameter_ref.element_name,
                new_value=float(value),
                line_index=parameter_ref.line_index,
                parameter_name=parameter_ref.parameter,
            )

            write_atp_file(elements, original_lines, execution_atp_path)
            _emit_event(
                event_callback,
                type="atp_written",
                message=f"ATP parametrizado salvo: {execution_atp_path.name}",
                run_index=run_index,
                total_runs=total_runs,
                value=float(value),
                progress=((run_index - 1) + 0.15) / total_runs,
                run_dir=run_dir,
                atp_path=execution_atp_path,
            )

            def _solver_status_callback(message: str) -> None:
                _emit_event(
                    event_callback,
                    type="solver_message",
                    message=message,
                    run_index=run_index,
                    total_runs=total_runs,
                    value=float(value),
                    progress=((run_index - 1) + 0.55) / total_runs,
                    run_dir=run_dir,
                )

            generated_lis_path = Path(
                solver_runner(
                    str(execution_atp_path),
                    timeout=solver_timeout,
                    status_callback=_solver_status_callback,
                )
            )

            _emit_event(
                event_callback,
                type="solver_finished",
                message=f"Solver finalizado para valor {value:g}",
                run_index=run_index,
                total_runs=total_runs,
                value=float(value),
                progress=((run_index - 1) + 0.8) / total_runs,
                run_dir=run_dir,
                lis_path=generated_lis_path,
            )

            relocated = _relocate_run_artifacts(
                base_atp_path=base_path,
                execution_atp_path=execution_atp_path,
                generated_lis_path=generated_lis_path,
                run_dir=run_dir,
            )
            result.atp_path = relocated["atp_path"]
            result.lis_path = relocated["lis_path"]

            _emit_event(
                event_callback,
                type="lis_ready",
                message=f"LIS pronto em {result.lis_path}",
                run_index=run_index,
                total_runs=total_runs,
                value=float(value),
                progress=((run_index - 1) + 0.9) / total_runs,
                run_dir=run_dir,
                lis_path=result.lis_path,
            )

            if lis_parser is not None and result.lis_path is not None:
                _emit_event(
                    event_callback,
                    type="lis_parsing",
                    message=f"Pos-processando {result.lis_path.name}",
                    run_index=run_index,
                    total_runs=total_runs,
                    value=float(value),
                    progress=((run_index - 1) + 0.95) / total_runs,
                    run_dir=run_dir,
                    lis_path=result.lis_path,
                )
                result.analysis = lis_parser(
                    result.lis_path,
                    run_dir,
                    float(value),
                    run_index,
                    total_runs,
                )

            result.status = "success"
            result.elapsed_seconds = time.monotonic() - run_started
            _emit_event(
                event_callback,
                type="run_succeeded",
                message=(
                    f"Run {run_index}/{total_runs} concluido em "
                    f"{result.elapsed_seconds:.2f}s"
                ),
                run_index=run_index,
                total_runs=total_runs,
                value=float(value),
                progress=run_index / total_runs,
                run_dir=run_dir,
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
                run_dir=run_dir,
            )
            result.atp_path = relocated.get("atp_path")
            result.lis_path = relocated.get("lis_path")

            _emit_event(
                event_callback,
                type="run_failed",
                message=(
                    f"Run {run_index}/{total_runs} falhou em "
                    f"{result.elapsed_seconds:.2f}s: {result.error}"
                ),
                run_index=run_index,
                total_runs=total_runs,
                value=float(value),
                progress=run_index / total_runs,
                run_dir=run_dir,
                error=result.error,
            )

            if not continue_on_error:
                summary.stopped_on_error = True
                break

    if cancel_event is not None and cancel_event.is_set():
        summary.cancelled = True

    summary.finished_at = datetime.now()

    if summary.cancelled:
        _emit_event(
            event_callback,
            type="sweep_cancelled",
            message="Sweep cancelado pelo usuario",
            run_index=len(summary.results),
            total_runs=total_runs,
            value=None,
            progress=min(1.0, len(summary.results) / max(1, total_runs)),
            output_dir=sweep_dir,
        )

    _emit_event(
        event_callback,
        type="sweep_finished",
        message=(
            f"Sweep finalizado: {summary.success_count} sucesso(s), "
            f"{summary.failure_count} falha(s), {summary.cancelled_count} cancelada(s)"
        ),
        run_index=len(summary.results),
        total_runs=total_runs,
        value=None,
        progress=1.0 if not summary.cancelled and not summary.stopped_on_error else min(
            1.0, len(summary.results) / max(1, total_runs)
        ),
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

    if generated_lis_path is not None and generated_lis_path.exists() and lis_target is not None:
        relocated_lis = _move_or_copy(generated_lis_path, lis_target)
        moved_sources.add(generated_lis_path.resolve())

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
