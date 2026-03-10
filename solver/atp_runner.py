from __future__ import annotations

import subprocess
import time
from pathlib import Path

ATP_EXECUTABLE = r"C:\ATP\tools\runATP.exe"


def run_atp_solver(atp_file_path: str, timeout: int = 600) -> str:
    """
    Executa uma simulação ATP usando runATP.exe e aguarda o término real da simulação.

    Detecção de término:
    - espera o .lis aparecer
    - monitora o final do .lis
    - detecta 'TOTAL ELAPSED TIME' ou 'JOB COMPLETED'
    """

    atp_path = Path(atp_file_path)

    if not atp_path.exists():
        raise FileNotFoundError(f"Arquivo .atp nao encontrado: {atp_file_path}")

    working_directory = atp_path.parent
    atp_name = atp_path.name
    base_name = atp_path.stem

    lis_lower = working_directory / f"{base_name}.lis"
    lis_upper = working_directory / f"{base_name}.LIS"

    print("\n===== ATP SIMULATION START =====")
    print("ATP executable:", ATP_EXECUTABLE)
    print("ATP input:", atp_path)
    print("Working directory:", working_directory)

    start_time = time.time()

    # executa runATP
    subprocess.Popen(
        [ATP_EXECUTABLE, atp_name],
        cwd=working_directory,
    )

    print("Solver iniciado, aguardando .lis...")

    # esperar .lis aparecer
    lis_path = None
    while True:

        if lis_lower.exists():
            lis_path = lis_lower
            break

        if lis_upper.exists():
            lis_path = lis_upper
            break

        if time.time() - start_time > timeout:
            raise TimeoutError("Timeout aguardando geracao do .lis")

        time.sleep(0.5)

    print("LIS detectado:", lis_path)

    # monitorar final da simulação
    print("Monitorando final da simulacao...")

    while True:

        if time.time() - start_time > timeout:
            raise TimeoutError("Timeout aguardando final da simulacao")

        try:
            with lis_path.open("rb") as f:

                f.seek(0, 2)
                size = f.tell()

                read_size = min(2000, size)

                f.seek(-read_size, 2)
                tail = f.read().decode(errors="ignore")

            if "TOTAL ELAPSED TIME" in tail or "JOB COMPLETED" in tail:
                break

        except Exception:
            pass

        time.sleep(1)

    elapsed = time.time() - start_time

    print("Simulacao concluida")
    print(f"Tempo total: {elapsed:.2f} s")
    print("===== ATP SIMULATION END =====\n")

    return str(lis_path)