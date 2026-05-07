from __future__ import annotations

import re
import subprocess
import threading
import time
from pathlib import Path
from typing import Callable, Optional

ATP_EXECUTABLE = r"C:\ATP\tools\runATP.exe"


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
                "ATP abortou (KILL) por formatacao de fim de linha invalida no .atp gerado "
                "(CR/LF inconsistente)."
                + (f"\nTrecho do LIS:\n{excerpt}" if excerpt else "")
            )
        return (
            "ATP abortou com EMTP error stop (KILL code) no arquivo .lis."
            + (f"\nTrecho do LIS:\n{excerpt}" if excerpt else "")
        )

    return None


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


def run_atp_solver(
    atp_file_path: str,
    timeout: int = 600,
    status_callback: Optional[Callable[[str], None]] = None,
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
        raise FileNotFoundError(f"Arquivo .atp nao encontrado: {atp_file_path}")

    missing_insert_dependencies = get_missing_insert_dependencies(atp_file_path)
    if missing_insert_dependencies:
        details = "\n".join([f"linha {line_no}: {target}" for line_no, target in missing_insert_dependencies])
        raise FileNotFoundError(
            "Arquivo(s) auxiliar(es) de $INSERT nao encontrado(s). "
            "Copie os includes do ATPDraw para a mesma pasta do .atp ou ajuste os caminhos.\n"
            f"{details}"
        )

    working_directory = atp_path.parent
    atp_name = atp_path.name
    base_name = atp_path.stem
    start_wall_time = time.time()
    start_monotonic = time.monotonic()

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

    _notify("===== ATP SIMULATION START =====")
    _notify(f"ATP executable: {ATP_EXECUTABLE}")
    _notify(f"ATP input: {atp_path}")
    _notify(f"Working directory: {working_directory}")

    # Executa runATP e aguarda o término real do processo.
    process = subprocess.Popen(
        [ATP_EXECUTABLE, atp_name],
        cwd=working_directory,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    _notify(f"Process started (pid={process.pid}). Waiting for completion...")

    # Alguns wrappers do ATP exibem "Hit any key to close this window" no fim.
    # Este thread garante fechamento automático sem depender de interação manual.
    auto_enter_thread = threading.Thread(target=_auto_press_enter, args=(process, status_callback), daemon=True)
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
                    _notify(f"[runATP] {clean}")
                    lower = clean.lower()
                    if "total execution time was" in lower or "atp finished at" in lower:
                        if not completion_event.is_set():
                            _notify("Completion marker detected in runATP output.")
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
    process_done = False
    process_done_at_monotonic = None

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

    try:
        while True:
            now_monotonic = time.monotonic()
            elapsed_total = now_monotonic - start_monotonic
            required_stable_window_sec = (
                stable_window_after_process_sec if process_done else stable_window_running_sec
            )
            lis_stable = _update_lis_state(required_stable_window_sec)

            if not process_done:
                polled = process.poll()
                if polled is not None:
                    return_code = polled
                    process_done = True
                    process_done_at_monotonic = now_monotonic
                    _notify(f"Process finished with return code {return_code}")
                    _notify("Waiting for LIS generation/stabilization...")

            # Se o wrapper travar no "Hit any key", finaliza automaticamente
            # quando houver marcador de conclusao e LIS estavel.
            if not process_done and completion_event.is_set() and lis_stable:
                _notify("Stable LIS + completion marker detected. Closing interactive runATP wrapper...")
                try:
                    process.terminate()
                    return_code = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    return_code = process.wait()

                process_done = True
                process_done_at_monotonic = now_monotonic
                _notify(f"Process finished with return code {return_code}")
                _notify("Waiting for LIS generation/stabilization...")

            # Fallback: alguns wrappers não emitem marcador de conclusão em execução parametrizada.
            # Se já existe LIS estável e sem output novo por alguns segundos, fecha o wrapper ocioso.
            if (
                not process_done
                and lis_stable
                and (now_monotonic - last_output_monotonic) > 8.0
                and elapsed_total > 12.0
            ):
                _notify("Stable LIS detected with idle wrapper output. Forcing wrapper shutdown...")
                try:
                    process.terminate()
                    return_code = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    return_code = process.wait()

                process_done = True
                process_done_at_monotonic = now_monotonic
                _notify(f"Process finished with return code {return_code}")
                _notify("Waiting for LIS generation/stabilization...")

            if process_done and lis_stable:
                break

            # Timeout global de execucao ATP.
            if elapsed_total > timeout:
                process.kill()
                process.wait()
                raise TimeoutError(f"Timeout aguardando termino do processo ATP ({timeout}s)")

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

        auto_enter_thread.join(timeout=1.0)
        output_thread.join(timeout=1.0)

    elapsed = time.monotonic() - start_monotonic

    _notify("Simulacao concluida")
    _notify(f"Tempo total: {elapsed:.2f} s")
    _notify("===== ATP SIMULATION END =====")

    if lis_path is None:
        raise RuntimeError(
            f"ATP process finished with code {return_code}, but .lis/.LIS was not generated"
        )

    lis_fatal_error = _detect_lis_fatal_error(lis_path)
    if lis_fatal_error is not None:
        raise RuntimeError(lis_fatal_error)

    if return_code != 0:
        _notify(f"Aviso: ATP finalizou com codigo {return_code}, mas gerou LIS valido.")

    _notify(f"LIS ready: {lis_path}")

    return str(lis_path)
