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
    - aguarda o término real do processo via process.wait()
    - valida existência de .lis/.LIS
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

    def _drain_output() -> None:
        if process.stdout is None:
            return
        try:
            for line in process.stdout:
                clean = line.strip()
                if clean:
                    _notify(f"[runATP] {clean}")
        except Exception:
            pass

    output_thread = threading.Thread(target=_drain_output, daemon=True)
    output_thread.start()

    try:
        return_code = process.wait(timeout=timeout)
        _notify(f"Process finished with return code {return_code}")
    except subprocess.TimeoutExpired as exc:
        process.kill()
        process.wait()
        raise TimeoutError(f"Timeout aguardando termino do processo ATP ({timeout}s)") from exc
    finally:
        try:
            if process.stdin is not None:
                process.stdin.close()
        except Exception:
            pass

        auto_enter_thread.join(timeout=1.0)
        output_thread.join(timeout=1.0)

    # Janela de tolerancia para flush final + estabilizacao do arquivo .lis.
    _notify("Waiting for LIS generation/stabilization...")
    grace_deadline = time.time() + 30.0
    stable_window_sec = 3.0
    lis_path = None
    last_size = None
    last_change_time = None

    while time.time() < grace_deadline:
        candidate = None
        if lis_lower.exists() and lis_lower.stat().st_size > 0:
            candidate = lis_lower
        elif lis_upper.exists() and lis_upper.stat().st_size > 0:
            candidate = lis_upper

        if candidate is not None:
            current_size = candidate.stat().st_size
            lis_path = candidate

            if last_size != current_size:
                last_size = current_size
                last_change_time = time.time()
            elif last_change_time is not None and (time.time() - last_change_time) >= stable_window_sec:
                break

        time.sleep(0.2)

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