from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from typing import Any, Callable, Optional

DEFAULT_ATP_ROOT = Path(r"C:\ATP")
ATP_EXECUTABLE = str(DEFAULT_ATP_ROOT / "tools" / "runATP.exe")
ATP_DIRECT_EXECUTABLE = DEFAULT_ATP_ROOT / "atpmingw" / "tpbig.exe"
ATP_DIRECT_SUPPORT_FILES = (
    "startup",
    "graphics",
    "graphics.aux",
    "graphics.std",
    "listsize.big",
)
LIS_TAIL_READ_BYTES = 64 * 1024
ATP_RESULT_STAGING_PREFIX = "lis_analysis_atp_result_"
ATP_WORKSPACE_PREFIX = "lis_analysis_atp_workspace_"


class ATPExecutionCancelled(RuntimeError):
    """Indica que a execucao ATP foi interrompida por solicitacao do usuario."""

    def __init__(self, message: str, lis_path: Path | None = None):
        super().__init__(message)
        self.lis_path = lis_path


def _windows_registry_atp_roots() -> list[Path]:
    if os.name != "nt":
        return []

    try:
        import winreg
    except ImportError:
        return []

    roots: list[Path] = []
    views = [
        getattr(winreg, "KEY_WOW64_32KEY", 0),
        getattr(winreg, "KEY_WOW64_64KEY", 0),
    ]
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in views:
            try:
                with winreg.OpenKey(
                    hive,
                    r"SOFTWARE\ATPINST",
                    0,
                    winreg.KEY_READ | view,
                ) as key:
                    value, _kind = winreg.QueryValueEx(key, "")
                if value:
                    roots.append(Path(str(value)))
            except OSError:
                continue
    return roots


def _normalize_atp_root(path: Path) -> Path:
    if path.name.lower() in {"atpmingw", "tools"}:
        return path.parent
    return path


def _candidate_atp_roots() -> list[Path]:
    candidates: list[Path] = []
    for variable in ("ATP_HOME", "ATPINST", "ATPDIR"):
        value = os.environ.get(variable)
        if value:
            candidates.append(Path(value))

    candidates.extend(_windows_registry_atp_roots())
    candidates.append(Path(ATP_EXECUTABLE).parent.parent)
    candidates.append(DEFAULT_ATP_ROOT)

    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        normalized = _normalize_atp_root(candidate)
        key = str(normalized).rstrip("\\/").lower()
        if key and key not in seen:
            seen.add(key)
            unique.append(normalized)
    return unique


def _discover_atp_executables(
    candidate_roots: Optional[list[Path]] = None,
) -> tuple[Optional[Path], Optional[Path]]:
    wrapper = None
    direct = None

    if candidate_roots is None:
        explicit_wrapper = os.environ.get("ATP_RUNATP")
        explicit_direct = os.environ.get("ATP_TPBIG")
        if explicit_wrapper and Path(explicit_wrapper).is_file():
            wrapper = Path(explicit_wrapper)
        if explicit_direct and Path(explicit_direct).is_file():
            direct = Path(explicit_direct)
        roots = _candidate_atp_roots()
    else:
        roots = [_normalize_atp_root(Path(root)) for root in candidate_roots]

    for root in roots:
        if wrapper is None:
            candidate = root / "tools" / "runATP.exe"
            if candidate.is_file():
                wrapper = candidate
        if direct is None:
            candidate = root / "atpmingw" / "tpbig.exe"
            if candidate.is_file():
                direct = candidate
        if wrapper is not None and direct is not None:
            break

    if wrapper is None:
        located = shutil.which("runATP.exe")
        if located:
            wrapper = Path(located)
    if direct is None:
        located = shutil.which("tpbig.exe")
        if located:
            direct = Path(located)

    return wrapper, direct

def validate_atp_executable_path(executable_path: str | Path) -> Path:
    """Valida um executável ATP escolhido manualmente."""
    candidate = Path(executable_path).expanduser()
    if not candidate.is_file():
        raise FileNotFoundError(f"Executável ATP não encontrado: {candidate}")

    if candidate.name.lower() not in {"tpbig.exe", "runatp.exe"}:
        raise ValueError(
            "Selecione o arquivo tpbig.exe ou runATP.exe da instalação do ATP."
        )
    return candidate.resolve()


def find_available_atp_executable() -> Path | None:
    """Retorna o melhor executável localizado automaticamente, se houver."""
    wrapper, direct = _discover_atp_executables()
    if direct is not None and all(
        (direct.parent / filename).is_file()
        for filename in ATP_DIRECT_SUPPORT_FILES
    ):
        return direct
    if wrapper is not None:
        return wrapper
    return direct


def _adaptive_poll_interval(idle_seconds: float) -> float:
    """Define o intervalo de polling com backoff progressivo em periodos ociosos."""
    if idle_seconds < 2.0:
        return 0.05
    if idle_seconds < 8.0:
        return 0.2
    return 0.5


def _extract_lis_error_excerpt(text: str, context_lines: int = 8) -> str:
    """Extrai um trecho curto ao redor de mensagens de erro criticas no LIS."""
    lines = text.splitlines()
    key_idx = None

    for idx, line in enumerate(lines):
        low = line.lower()
        if "kill code" in low or "emtp error stop" in low or "temporary error stop" in low:
            key_idx = idx
            break

    if key_idx is None:
        return ""

    start = max(0, key_idx - 1)
    end = min(len(lines), key_idx + context_lines)
    excerpt_lines = [ln.rstrip() for ln in lines[start:end] if ln.strip()]
    return "\n".join(excerpt_lines)


def _detect_lis_fatal_error(lis_path: Path) -> Optional[str]:
    """Retorna mensagem de erro fatal detectada no .lis, ou None se não houver."""
    try:
        text = lis_path.read_text(errors="replace")
    except Exception:
        return None

    lower_text = text.lower()

    if "error during connection of disk file of $insert" in lower_text or "halt in insert" in lower_text:
        include_line = None
        for line in text.splitlines():
            if "$insert" in line.lower():
                include_line = line.strip()
                if include_line:
                    break

        details = f" Include problemático: {include_line}" if include_line else ""
        return (
            "ATP interrompeu na diretiva $INSERT (arquivo de include não encontrado ou inacessível)."
            + details
        )

    if "temporary error stop" in lower_text:
        excerpt = _extract_lis_error_excerpt(text)
        return (
            "ATP reportou 'Temporary error stop' no arquivo .lis."
            + (f"\nTrecho do LIS:\n{excerpt}" if excerpt else "")
        )

    if "emtp error stop" in lower_text or "kill code" in lower_text:
        excerpt = _extract_lis_error_excerpt(text)
        if "carriage return" in lower_text and "line feed" in lower_text:
            return (
                "ATP abortou (KILL) por formatação de fim de linha inválida no .atp gerado "
                "(CR/LF inconsistente)."
                + (f"\nTrecho do LIS:\n{excerpt}" if excerpt else "")
            )
        return (
            "ATP abortou com EMTP error stop (KILL code) no arquivo .lis."
            + (f"\nTrecho do LIS:\n{excerpt}" if excerpt else "")
        )

    return None


def _lis_has_completion_marker(lis_path: Path) -> bool:
    """Confirma que o ATP escreveu o bloco final de temporizacao do LIS."""
    try:
        with lis_path.open("rb") as stream:
            stream.seek(0, 2)
            size = stream.tell()
            stream.seek(max(0, size - LIS_TAIL_READ_BYTES))
            tail = stream.read().decode(errors="replace").lower()
    except Exception:
        return False

    has_list_sizes = "actual list sizes for the preceding solution follow" in tail
    has_timing_end = "seconds after deltat-loop" in tail and re.search(
        r"^\s*totals\s*:", tail, flags=re.MULTILINE
    ) is not None
    return has_list_sizes and has_timing_end


ATP_FLOAT_PATTERN = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?"


def _extract_lis_simulation_progress(text: str) -> tuple[float, str] | None:
    """Extrai progresso real do tempo simulado ou das rodadas estatisticas."""
    if not text.strip():
        return None

    statistical_total_matches = list(
        re.finditer(
            r"\bNENERG\s*=\s*(\d+)\s+simulations?\b",
            text,
            flags=re.IGNORECASE,
        )
    )
    planned_total_match = re.search(
        r"^Misc\.\s+data\.\s+(?:\d+\s+){8}(\d+)",
        text,
        flags=re.MULTILINE,
    )
    total = 0
    if statistical_total_matches:
        total = int(statistical_total_matches[-1].group(1))
    elif planned_total_match is not None:
        total = int(planned_total_match.group(1))

    if total > 0:
        current_matches = list(
            re.finditer(
                r"simulation\s+number\s+(\d+)",
                text,
                flags=re.IGNORECASE,
            )
        )
        current = int(current_matches[-1].group(1)) if current_matches else 0
        progress = max(0.0, min(1.0, current / total))
        if current == 0:
            return progress, f"Preparando {total} casos estat\u00edsticos"
        return progress, f"Simulando caso {current}/{total}"

    misc_pattern = re.compile(
        rf"Misc\.\s+data\.\s+({ATP_FLOAT_PATTERN})\s+({ATP_FLOAT_PATTERN})",
        flags=re.IGNORECASE,
    )
    misc_match = misc_pattern.search(text)
    if misc_match is None:
        return None

    try:
        tmax = float(misc_match.group(2))
    except ValueError:
        return None
    if tmax <= 0:
        return None

    simulated_times: list[float] = []
    step_pattern = re.compile(
        rf"^\s*\d+\s+({ATP_FLOAT_PATTERN})(?:\s|$)",
        flags=re.MULTILINE,
    )
    for match in step_pattern.finditer(text):
        try:
            simulated_times.append(float(match.group(1)))
        except ValueError:
            continue

    plot_span_pattern = re.compile(
        rf"Plot\s+timespan.*?=\s*{ATP_FLOAT_PATTERN}\s+({ATP_FLOAT_PATTERN})",
        flags=re.IGNORECASE,
    )
    for match in plot_span_pattern.finditer(text):
        try:
            simulated_times.append(float(match.group(1)))
        except ValueError:
            continue

    if not simulated_times:
        return 0.0, "Iniciando passos de tempo"

    current_time = max(0.0, max(simulated_times))
    progress = max(0.0, min(1.0, current_time / tmax))
    return progress, f"Simulando t={current_time:g}s de {tmax:g}s"


def _read_lis_progress_text(lis_path: Path) -> str:
    # Le cabecalho e cauda do LIS sem reler arquivos grandes por inteiro.
    header_bytes = 64 * 1024
    tail_bytes = 256 * 1024
    with lis_path.open("rb") as stream:
        stream.seek(0, 2)
        size = stream.tell()
        if size <= header_bytes + tail_bytes:
            stream.seek(0)
            data = stream.read()
        else:
            stream.seek(0)
            header = stream.read(header_bytes)
            stream.seek(-tail_bytes, 2)
            data = header + b"\n" + stream.read(tail_bytes)
    return data.decode(errors="replace")


def _stage_direct_solver_support(
    working_directory: Path,
    support_directory: Path,
) -> list[Path]:
    """Copia apenas suportes ausentes e retorna os arquivos que devem ser removidos."""
    copied: list[Path] = []
    try:
        for filename in ATP_DIRECT_SUPPORT_FILES:
            source = support_directory / filename
            if not source.is_file():
                raise FileNotFoundError(f"Suporte ATP ausente: {source}")

            destination = working_directory / filename
            if destination.exists():
                continue

            shutil.copy2(source, destination)
            copied.append(destination)
    except Exception:
        _cleanup_direct_solver_support(copied)
        raise
    return copied


def _cleanup_direct_solver_support(copied_files: list[Path]) -> None:
    for path in reversed(copied_files):
        try:
            path.unlink()
        except OSError:
            pass


def _snapshot_matching_files(
    directory: Path,
    predicate: Callable[[Path], bool],
) -> dict[str, tuple[int, int]]:
    snapshot: dict[str, tuple[int, int]] = {}
    for candidate in directory.iterdir():
        if not candidate.is_file() or not predicate(candidate):
            continue
        try:
            stat = candidate.stat()
            snapshot[str(candidate.resolve()).lower()] = (
                int(stat.st_mtime_ns),
                int(stat.st_size),
            )
        except OSError:
            continue
    return snapshot


def _is_atp_temporary_file(path: Path) -> bool:
    name = path.name.lower()
    return (
        (name.startswith("dum") and name.endswith(".bin"))
        or path.suffix.lower() in {".tmp", ".temp"}
    )


def _is_cancelled_atp_result(path: Path) -> bool:
    return path.suffix.lower() in {".lis", ".pl4", ".dbg"}


def _remove_files_changed_since_snapshot(
    directory: Path,
    snapshot: dict[str, tuple[int, int]],
    predicate: Callable[[Path], bool],
) -> list[Path]:
    removed: list[Path] = []
    for candidate in directory.iterdir():
        if not candidate.is_file() or not predicate(candidate):
            continue
        try:
            stat = candidate.stat()
            current = (int(stat.st_mtime_ns), int(stat.st_size))
            previous = snapshot.get(str(candidate.resolve()).lower())
            if previous == current:
                continue
            candidate.unlink()
            removed.append(candidate)
        except OSError:
            continue
    return removed


def _parse_insert_targets(atp_path: Path) -> list[tuple[int, str]]:
    """Extrai alvos de diretivas $INSERT no formato (linha, caminho)."""
    targets: list[tuple[int, str]] = []
    for idx, raw_line in enumerate(atp_path.read_text(errors="replace").splitlines(), start=1):
        stripped = raw_line.strip()
        if not stripped.upper().startswith("$INSERT"):
            continue

        parts = stripped.split(",", 1)
        if len(parts) < 2:
            continue

        target = parts[1].strip().strip('"').strip("'")
        if target:
            targets.append((idx, target))
    return targets


def _resolve_insert_target_path(working_directory: Path, target: str) -> Optional[Path]:
    """Resolve o caminho de um include $INSERT para um arquivo existente."""
    is_windows_absolute = re.match(r"^[a-zA-Z]:[\\/]", target) is not None
    if is_windows_absolute:
        win_path = Path(target)
        return win_path if win_path.exists() else None

    normalized_target = target.replace("\\", "/")
    candidate = Path(normalized_target)
    if candidate.is_absolute():
        return candidate if candidate.exists() else None

    local_candidate = (working_directory / candidate).resolve()
    return local_candidate if local_candidate.exists() else None


def get_missing_insert_dependencies(atp_file_path: str) -> list[tuple[int, str]]:
    """Retorna lista de includes $INSERT faltantes como (linha, caminho)."""
    atp_path = Path(atp_file_path)
    if not atp_path.exists():
        return []

    missing: list[tuple[int, str]] = []
    for line_no, target in _parse_insert_targets(atp_path):
        resolved = _resolve_insert_target_path(atp_path.parent, target)
        if resolved is None:
            missing.append((line_no, target))
    return missing


def _auto_press_enter(process: subprocess.Popen, status_callback: Optional[Callable[[str], None]] = None) -> None:
    """Envia ENTER periodicamente para destravar prompts interativos do runATP."""
    while process.poll() is None:
        try:
            if process.stdin is not None:
                process.stdin.write("\n")
                process.stdin.flush()
        except Exception:
            break
        time.sleep(1.0)


def _terminate_process_tree(process: subprocess.Popen) -> None:
    """Encerra o solver e, no Windows, eventuais processos filhos do wrapper."""
    if process.poll() is not None:
        return

    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5,
                check=False,
            )
        except Exception:
            pass

    if process.poll() is not None:
        return

    try:
        process.terminate()
    except Exception:
        pass

    try:
        process.wait(timeout=3)
    except subprocess.TimeoutExpired:
        try:
            process.kill()
            process.wait(timeout=3)
        except Exception:
            pass
    except Exception:
        pass


def _run_atp_solver_in_workspace(
    atp_file_path: str,
    timeout: int = 600,
    status_callback: Optional[Callable[[str], None]] = None,
    cancel_event: Any | None = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    atp_executable_path: str | Path | None = None,
) -> str:
    """
    Executa uma simulação ATP usando runATP.exe e aguarda o término real da simulação.

    Detecção de término:
    - inicia o solver em processo separado
    - detecta conclusão pelo processo ou por marcador no output
    - valida existência/estabilização de .lis/.LIS
    """

    atp_path = Path(atp_file_path)

    if not atp_path.exists():
        raise FileNotFoundError(f"Arquivo .atp não encontrado: {atp_file_path}")

    missing_insert_dependencies = get_missing_insert_dependencies(atp_file_path)
    if missing_insert_dependencies:
        details = "\n".join([f"linha {line_no}: {target}" for line_no, target in missing_insert_dependencies])
        raise FileNotFoundError(
            "Arquivo(s) auxiliar(es) de $INSERT não encontrado(s). "
            "Copie os includes do ATPDraw para a mesma pasta do .atp ou ajuste os caminhos.\n"
            f"{details}"
        )

    working_directory = atp_path.parent
    atp_name = atp_path.name
    base_name = atp_path.stem
    start_wall_time = time.time()
    start_monotonic = time.monotonic()
    temporary_snapshot = _snapshot_matching_files(
        working_directory,
        _is_atp_temporary_file,
    )
    def _is_current_atp_result(path: Path) -> bool:
        return (
            _is_cancelled_atp_result(path)
            and path.stem.lower() == base_name.lower()
        )

    result_snapshot = _snapshot_matching_files(
        working_directory,
        _is_current_atp_result,
    )

    lis_snapshot: dict[str, tuple[int, int]] = {}
    for existing in list(working_directory.glob("*.lis")) + list(working_directory.glob("*.LIS")):
        try:
            st = existing.stat()
            lis_snapshot[str(existing.resolve()).lower()] = (int(st.st_mtime_ns), int(st.st_size))
        except Exception:
            continue

    lis_lower = working_directory / f"{base_name}.lis"
    lis_upper = working_directory / f"{base_name}.LIS"
    start_floor = start_wall_time - 1.0

    def _is_recent_lis(path: Path) -> bool:
        try:
            st = path.stat()
        except Exception:
            return False

        if st.st_size <= 0:
            return False

        key = str(path.resolve()).lower()
        previous = lis_snapshot.get(key)
        if previous is None:
            # Arquivo não existia no snapshot pré-run: é novo desta execução,
            # mesmo que metadados de mtime venham com valor antigo.
            return True

        prev_mtime_ns, prev_size = previous
        if int(st.st_mtime_ns) > prev_mtime_ns:
            return True
        if int(st.st_size) != prev_size:
            return True
        return False

    def _discover_recent_lis() -> Optional[Path]:
        """Procura LIS recente quando ATP gerar nome diferente do .atp executado."""
        candidates = list(working_directory.glob("*.lis")) + list(working_directory.glob("*.LIS"))
        filtered: list[Path] = []
        for p in candidates:
            if _is_recent_lis(p):
                filtered.append(p)

        if not filtered:
            return None

        filtered.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return filtered[0]

    def _notify(message: str) -> None:
        print(message)
        if status_callback is not None:
            try:
                status_callback(message)
            except Exception:
                pass

    _notify("===== INÍCIO DA SIMULAÇÃO ATP =====")
    selected_executable = None
    if atp_executable_path:
        selected_executable = validate_atp_executable_path(atp_executable_path)
        if selected_executable.name.lower() == "tpbig.exe":
            wrapper_executable, direct_executable = None, selected_executable
        else:
            wrapper_executable, direct_executable = selected_executable, None
        _notify(f"Executável ATP selecionado manualmente: {selected_executable}")
    else:
        wrapper_executable, direct_executable = _discover_atp_executables()

    direct_support_files: list[Path] = []
    use_direct_solver = False
    if direct_executable is not None:
        try:
            direct_support_files = _stage_direct_solver_support(
                working_directory,
                direct_executable.parent,
            )
            use_direct_solver = True
        except Exception as exc:
            if selected_executable is not None:
                raise RuntimeError(
                    f"Não foi possível usar o tpbig.exe selecionado: {exc}"
                ) from exc
            _notify(f"Solver ATP direto indisponível ({exc}). Usando runATP.exe.")

    # Mantem compatibilidade com integracoes que interceptam o processo e
    # transforma a falha real ao iniciar em uma mensagem de configuracao clara.
    if wrapper_executable is None:
        wrapper_executable = Path(ATP_EXECUTABLE)

    solver_command = [str(wrapper_executable), atp_name]
    if use_direct_solver:
        solver_command = [
            str(direct_executable),
            "DISK",
            atp_name,
            "s",
            "-R",
        ]

    _notify(f"Modo ATP: {'tpbig direto' if use_direct_solver else 'intermediado por runATP'}")
    _notify(f"Executável ATP: {solver_command[0]}")
    _notify(f"Arquivo ATP: {atp_path}")
    _notify(f"Diretório de trabalho: {working_directory}")

    # Executa o solver direto quando disponivel; mantem o wrapper como fallback.
    try:
        process = subprocess.Popen(
            solver_command,
            cwd=working_directory,
            stdin=subprocess.DEVNULL if use_direct_solver else subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
    except FileNotFoundError as exc:
        _cleanup_direct_solver_support(direct_support_files)
        raise FileNotFoundError(
            "Executável ATP não encontrado. Instale o ATP ou configure ATP_HOME, "
            "ATP_TPBIG ou ATP_RUNATP."
        ) from exc
    except Exception:
        _cleanup_direct_solver_support(direct_support_files)
        raise
    _notify(f"Processo iniciado (pid={process.pid}). Aguardando conclusão...")

    # Alguns wrappers do ATP exibem "Hit any key to close this window" no fim.
    auto_enter_thread = None
    if not use_direct_solver:
        auto_enter_thread = threading.Thread(
            target=_auto_press_enter,
            args=(process, status_callback),
            daemon=True,
        )
        auto_enter_thread.start()

    completion_event = threading.Event()
    last_output_monotonic = time.monotonic()

    def _drain_output() -> None:
        nonlocal last_output_monotonic
        if process.stdout is None:
            return
        try:
            for line in process.stdout:
                clean = line.strip()
                if clean:
                    last_output_monotonic = time.monotonic()
                    lower = clean.lower()
                    if not use_direct_solver:
                        _notify(f"[runATP] {clean}")
                    if "total execution time was" in lower or "atp finished at" in lower:
                        if not completion_event.is_set():
                            _notify("Marcador de conclusão detectado na saída do runATP.")
                        completion_event.set()
        except Exception:
            pass

    output_thread = threading.Thread(target=_drain_output, daemon=True)
    output_thread.start()

    return_code = None
    lis_path = None
    last_size = None
    last_change_monotonic = None
    stable_window_running_sec = 3.0
    stable_window_after_process_sec = 1.0
    stable_window_with_completion_marker_sec = 0.75
    process_done = False
    process_done_at_monotonic = None
    last_progress_check_monotonic = 0.0
    last_reported_progress = -1.0
    last_progress_detail = ""

    def _report_lis_progress(*, force: bool = False) -> None:
        nonlocal last_progress_check_monotonic, last_reported_progress, last_progress_detail
        if progress_callback is None or lis_path is None:
            return

        now = time.monotonic()
        if not force and (now - last_progress_check_monotonic) < 0.25:
            return
        last_progress_check_monotonic = now

        try:
            progress_state = _extract_lis_simulation_progress(
                _read_lis_progress_text(lis_path)
            )
        except OSError:
            return
        if progress_state is None:
            return

        progress, detail = progress_state
        should_report = (
            force
            or progress >= 1.0
            or progress - last_reported_progress >= 0.005
            or detail != last_progress_detail
        )
        if not should_report:
            return

        last_reported_progress = max(last_reported_progress, progress)
        last_progress_detail = detail
        try:
            progress_callback(last_reported_progress, detail)
        except Exception:
            pass

    def _update_lis_state(required_stable_window_sec: float) -> bool:
        nonlocal lis_path, last_size, last_change_monotonic

        candidate = None
        try:
            if lis_lower.exists() and _is_recent_lis(lis_lower):
                candidate = lis_lower
            elif lis_upper.exists() and _is_recent_lis(lis_upper):
                candidate = lis_upper
            else:
                candidate = _discover_recent_lis()
        except Exception:
            candidate = None

        if candidate is None:
            return False

        lis_path = candidate
        try:
            current_size = candidate.stat().st_size
        except Exception:
            return False

        if last_size != current_size:
            last_size = current_size
            last_change_monotonic = time.monotonic()
            return False

        return (
            last_change_monotonic is not None
            and (time.monotonic() - last_change_monotonic) >= required_stable_window_sec
        )

    def _cleanup_cancelled_results() -> list[Path]:
        return _remove_files_changed_since_snapshot(
            working_directory,
            result_snapshot,
            _is_current_atp_result,
        )

    try:
        while True:
            if cancel_event is not None and cancel_event.is_set():
                _notify("Cancelamento solicitado. Encerrando o processo ATP...")
                _update_lis_state(0.0)
                _terminate_process_tree(process)
                removed_results = _cleanup_cancelled_results()
                if removed_results:
                    _notify(
                        f"Removido(s) {len(removed_results)} resultado(s) ATP incompleto(s)."
                    )
                raise ATPExecutionCancelled("Simulação ATP cancelada pelo usuário")

            now_monotonic = time.monotonic()
            elapsed_total = now_monotonic - start_monotonic
            required_stable_window_sec = (
                stable_window_after_process_sec if process_done else stable_window_running_sec
            )
            lis_stable = _update_lis_state(required_stable_window_sec)
            _report_lis_progress()
            marker_check_ready = bool(
                lis_path is not None
                and last_change_monotonic is not None
                and (now_monotonic - last_change_monotonic) >= 0.25
            )
            lis_completion_marker = bool(
                marker_check_ready and _lis_has_completion_marker(lis_path)
            )
            if (
                not lis_stable
                and lis_completion_marker
                and last_change_monotonic is not None
                and (now_monotonic - last_change_monotonic)
                >= stable_window_with_completion_marker_sec
            ):
                lis_stable = True

            if not process_done:
                polled = process.poll()
                if polled is not None:
                    return_code = polled
                    process_done = True
                    process_done_at_monotonic = now_monotonic
                    _notify(f"Processo finalizado com código de retorno {return_code}")
                    _notify("Aguardando geração e estabilização do LIS...")

            # Se o wrapper travar no "Hit any key", finaliza automaticamente
            # quando houver marcador de conclusao e LIS estavel.
            if (
                not use_direct_solver
                and not process_done
                and completion_event.is_set()
                and lis_stable
            ):
                _notify("LIS estável e marcador de conclusão detectados. Encerrando a janela interativa do runATP...")
                try:
                    process.terminate()
                    return_code = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    return_code = process.wait()

                process_done = True
                process_done_at_monotonic = now_monotonic
                _notify(f"Processo finalizado com código de retorno {return_code}")
                _notify("Aguardando geração e estabilização do LIS...")

            # O bloco final do LIS e o tamanho estavel confirmam a conclusao
            # sem depender dos 8 segundos de ociosidade do wrapper interativo.
            if (
                not use_direct_solver
                and not process_done
                and lis_completion_marker
                and lis_stable
            ):
                _notify("LIS completo e estável detectado. Encerrando a janela interativa do runATP...")
                try:
                    process.terminate()
                    return_code = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    return_code = process.wait()

                process_done = True
                process_done_at_monotonic = now_monotonic
                _notify(f"Processo finalizado com código de retorno {return_code}")
                _notify("Aguardando geração e estabilização do LIS...")

            # Fallback: alguns wrappers não emitem marcador de conclusão em execução parametrizada.
            # Se já existe LIS estável e sem output novo por alguns segundos, fecha o wrapper ocioso.
            if (
                not use_direct_solver
                and not process_done
                and lis_stable
                and (now_monotonic - last_output_monotonic) > 8.0
                and elapsed_total > 12.0
            ):
                _notify("LIS estável detectado sem nova saída do runATP. Encerrando o intermediador...")
                try:
                    process.terminate()
                    return_code = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    return_code = process.wait()

                process_done = True
                process_done_at_monotonic = now_monotonic
                _notify(f"Processo finalizado com código de retorno {return_code}")
                _notify("Aguardando geração e estabilização do LIS...")

            if process_done and lis_stable:
                break

            # Timeout global de execucao ATP.
            if elapsed_total > timeout:
                process.kill()
                process.wait()
                raise TimeoutError(f"Tempo limite excedido ao aguardar o término do processo ATP ({timeout}s)")

            # Tolerancia curta para flush apos processo terminar.
            if (
                process_done
                and process_done_at_monotonic is not None
                and (now_monotonic - process_done_at_monotonic) > 30.0
            ):
                break

            last_activity_monotonic = last_output_monotonic
            if last_change_monotonic is not None:
                last_activity_monotonic = max(last_activity_monotonic, last_change_monotonic)
            if process_done_at_monotonic is not None:
                last_activity_monotonic = max(last_activity_monotonic, process_done_at_monotonic)

            idle_for = max(0.0, now_monotonic - last_activity_monotonic)
            poll_interval = _adaptive_poll_interval(idle_for)

            remaining_timeout = max(0.0, float(timeout) - elapsed_total)
            if remaining_timeout > 0:
                poll_interval = min(poll_interval, max(0.01, remaining_timeout))

            time.sleep(poll_interval)
    finally:
        try:
            if process.stdin is not None:
                process.stdin.close()
        except Exception:
            pass

        if auto_enter_thread is not None:
            auto_enter_thread.join(timeout=1.0)
        output_thread.join(timeout=1.0)
        _cleanup_direct_solver_support(direct_support_files)
        removed_temporary = _remove_files_changed_since_snapshot(
            working_directory,
            temporary_snapshot,
            _is_atp_temporary_file,
        )
        if removed_temporary:
            _notify(f"Removido(s) {len(removed_temporary)} arquivo(s) temporário(s) do ATP.")

    if cancel_event is not None and cancel_event.is_set():
        removed_results = _cleanup_cancelled_results()
        if removed_results:
            _notify(f"Removido(s) {len(removed_results)} resultado(s) ATP incompleto(s).")
        raise ATPExecutionCancelled("Simulação ATP cancelada pelo usuário")

    elapsed = time.monotonic() - start_monotonic

    _notify("Simulação concluída")
    _notify(f"Tempo total: {elapsed:.2f} s")
    _notify("===== FIM DA SIMULAÇÃO ATP =====")

    if lis_path is None:
        raise RuntimeError(
            f"O processo ATP terminou com o código {return_code}, mas não gerou um arquivo .lis/.LIS"
        )

    lis_fatal_error = _detect_lis_fatal_error(lis_path)
    if lis_fatal_error is not None:
        raise RuntimeError(lis_fatal_error)

    if return_code != 0:
        _notify(f"Aviso: ATP finalizou com código {return_code}, mas gerou LIS válido.")

    _report_lis_progress(force=True)
    if progress_callback is not None:
        try:
            progress_callback(1.0, "Simula\u00e7\u00e3o ATP conclu\u00edda")
        except Exception:
            pass
    _notify(f"LIS pronto: {lis_path}")

    return str(lis_path)


def _isolated_workspace_destination(workspace_root: Path, source: Path) -> Path:
    """Mapeia um caminho absoluto sem alterar a estrutura usada por includes relativos."""
    resolved = source.resolve()
    anchor_label = re.sub(r"[^A-Za-z0-9_.-]+", "_", resolved.anchor).strip("_")
    if not anchor_label:
        anchor_label = "root"

    relative_parts = resolved.parts[1:] if resolved.anchor else resolved.parts
    return workspace_root / "files" / anchor_label / Path(*relative_parts)


def _prepare_isolated_workspace(atp_path: Path) -> tuple[Path, Path]:
    """Copia o ATP e todos os $INSERT relativos para um workspace exclusivo."""
    source_atp = atp_path.resolve()
    workspace_root = Path(tempfile.mkdtemp(prefix=ATP_WORKSPACE_PREFIX))
    pending = [source_atp]
    copied: dict[Path, Path] = {}

    try:
        while pending:
            source = pending.pop()
            source = source.resolve()
            if source in copied:
                continue
            if not source.is_file():
                raise FileNotFoundError(f"Dependência ATP não encontrada: {source}")

            destination = _isolated_workspace_destination(workspace_root, source)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
            copied[source] = destination

            for line_no, target in _parse_insert_targets(source):
                normalized = target.replace("\\", "/")
                is_windows_absolute = re.match(r"^[A-Za-z]:[\\/]", target) is not None
                target_path = Path(normalized)

                # Includes absolutos permanecem apontando para a instalacao/projeto
                # original. Eles sao somente leitura para o ATP.
                if is_windows_absolute or target_path.is_absolute():
                    resolved = _resolve_insert_target_path(source.parent, target)
                    if resolved is None:
                        raise FileNotFoundError(
                            f"$INSERT não encontrado em {source.name}, linha {line_no}: {target}"
                        )
                    continue

                dependency = (source.parent / target_path).resolve()
                if not dependency.is_file():
                    raise FileNotFoundError(
                        f"$INSERT não encontrado em {source.name}, linha {line_no}: {target}"
                    )
                pending.append(dependency)

        return workspace_root, copied[source_atp]
    except Exception:
        shutil.rmtree(workspace_root, ignore_errors=True)
        raise


def _stage_atp_results(executed_atp: Path, generated_lis: Path) -> Path:
    """Preserva resultados validos fora do workspace antes de remove-lo."""
    staging_dir = Path(tempfile.mkdtemp(prefix=ATP_RESULT_STAGING_PREFIX))
    try:
        staged_lis = staging_dir / generated_lis.name
        shutil.copy2(generated_lis, staged_lis)

        stems = {executed_atp.stem.lower(), generated_lis.stem.lower()}
        for candidate in executed_atp.parent.iterdir():
            if (
                candidate.is_file()
                and candidate.stem.lower() in stems
                and candidate.suffix.lower() in {".pl4", ".dbg"}
            ):
                shutil.copy2(candidate, staging_dir / candidate.name)
        return staged_lis
    except Exception:
        shutil.rmtree(staging_dir, ignore_errors=True)
        raise


def is_staged_atp_result(path: Path | str) -> bool:
    candidate = Path(path)
    staging_parent = candidate.parent
    try:
        system_temp = Path(tempfile.gettempdir()).resolve()
        return (
            staging_parent.name.startswith(ATP_RESULT_STAGING_PREFIX)
            and staging_parent.resolve().parent == system_temp
        )
    except OSError:
        return False


def iter_staged_atp_artifacts(lis_path: Path | str) -> tuple[Path, ...]:
    """Lista LIS e sidecars pertencentes ao mesmo resultado temporariamente preservado."""
    candidate = Path(lis_path)
    if not is_staged_atp_result(candidate) or not candidate.parent.is_dir():
        return (candidate,) if candidate.exists() else ()
    return tuple(path for path in candidate.parent.iterdir() if path.is_file())


def cleanup_staged_atp_result(path: Path | str | None) -> None:
    """Remove com seguranca apenas o diretorio de staging criado por este modulo."""
    if path is None:
        return
    candidate = Path(path)
    if is_staged_atp_result(candidate):
        shutil.rmtree(candidate.parent, ignore_errors=True)


def run_atp_solver(
    atp_file_path: str,
    timeout: int = 600,
    status_callback: Optional[Callable[[str], None]] = None,
    cancel_event: Any | None = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    atp_executable_path: str | Path | None = None,
) -> str:
    """Executa o ATP em workspace temporario e devolve apenas resultados validos."""
    source_atp = Path(atp_file_path)
    if not source_atp.is_file():
        raise FileNotFoundError(f"Arquivo .atp não encontrado: {atp_file_path}")

    def notify(message: str) -> None:
        if status_callback is not None:
            try:
                status_callback(message)
            except Exception:
                pass

    workspace_root: Path | None = None
    try:
        notify("Preparando workspace isolado da simulacao ATP...")
        workspace_root, isolated_atp = _prepare_isolated_workspace(source_atp)
        notify(f"Workspace isolado pronto: {workspace_root}")

        generated_lis = Path(
            _run_atp_solver_in_workspace(
                str(isolated_atp),
                timeout=timeout,
                status_callback=status_callback,
                cancel_event=cancel_event,
                progress_callback=progress_callback,
                atp_executable_path=atp_executable_path,
            )
        )
        staged_lis = _stage_atp_results(isolated_atp, generated_lis)
        notify("Resultados validos preservados; limpando workspace isolado...")
        return str(staged_lis)
    finally:
        if workspace_root is not None:
            shutil.rmtree(workspace_root, ignore_errors=True)
