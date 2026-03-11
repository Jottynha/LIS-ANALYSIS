from __future__ import annotations

import subprocess
import threading
import time
from pathlib import Path
from typing import Callable, Optional

ATP_EXECUTABLE = r"C:\ATP\tools\runATP.exe"


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

    working_directory = atp_path.parent
    atp_name = atp_path.name
    base_name = atp_path.stem

    lis_lower = working_directory / f"{base_name}.lis"
    lis_upper = working_directory / f"{base_name}.LIS"

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

    start_time = time.time()

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

    def _drain_output() -> None:
        if process.stdout is None:
            return
        try:
            for line in process.stdout:
                clean = line.strip()
                if clean:
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
    last_change_time = None
    stable_window_sec = 3.0
    process_done = False
    process_done_at = None

    def _update_lis_state() -> bool:
        nonlocal lis_path, last_size, last_change_time

        candidate = None
        try:
            if lis_lower.exists() and lis_lower.stat().st_size > 0:
                candidate = lis_lower
            elif lis_upper.exists() and lis_upper.stat().st_size > 0:
                candidate = lis_upper
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
            last_change_time = time.time()
            return False

        return last_change_time is not None and (time.time() - last_change_time) >= stable_window_sec

    try:
        while True:
            now = time.time()
            elapsed_total = now - start_time
            lis_stable = _update_lis_state()

            if not process_done:
                polled = process.poll()
                if polled is not None:
                    return_code = polled
                    process_done = True
                    process_done_at = now
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
                process_done_at = now
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
            if process_done and process_done_at is not None and (now - process_done_at) > 30.0:
                break

            time.sleep(0.2)
    finally:
        try:
            if process.stdin is not None:
                process.stdin.close()
        except Exception:
            pass

        auto_enter_thread.join(timeout=1.0)
        output_thread.join(timeout=1.0)

    elapsed = time.time() - start_time

    _notify("Simulacao concluida")
    _notify(f"Tempo total: {elapsed:.2f} s")
    _notify("===== ATP SIMULATION END =====")

    if lis_path is None:
        raise RuntimeError(
            f"ATP process finished with code {return_code}, but .lis/.LIS was not generated"
        )

    if return_code != 0:
        _notify(f"Aviso: ATP finalizou com codigo {return_code}, mas gerou LIS valido.")

    _notify(f"LIS ready: {lis_path}")

    return str(lis_path)