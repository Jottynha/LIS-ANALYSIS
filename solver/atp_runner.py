from __future__ import annotations

import subprocess
import time
from pathlib import Path

ATP_EXECUTABLE = r"C:\ATP\tools\runATP.exe"


def run_atp_solver(atp_file_path: str, timeout: int = 600) -> str:
    """
    Executa o solver ATP via runATP.exe.

    Fluxo:
    1. Executa runATP.exe
    2. Espera o .lis aparecer
    3. Espera o .lis parar de crescer
    4. Verifica se a simulação terminou corretamente

    Args:
        atp_file_path: caminho para o arquivo .atp
        timeout: tempo máximo total (segundos)

    Returns:
        Caminho para o arquivo .lis gerado
    """

    atp_path = Path(atp_file_path)

    if not atp_path.exists():
        raise FileNotFoundError(f"Arquivo .atp não encontrado: {atp_file_path}")

    working_directory = atp_path.parent
    atp_name = atp_path.name
    base_name = atp_path.stem

    lis_lower = working_directory / f"{base_name}.lis"
    lis_upper = working_directory / f"{base_name}.LIS"

    print("\n===== ATP SIMULATION START =====")
    print("Executable:", ATP_EXECUTABLE)
    print("Input file:", atp_path)
    print("Working directory:", working_directory)

    start_time = time.time()

    # -------------------------
    # Executa runATP
    # -------------------------

    process = subprocess.Popen(
        [ATP_EXECUTABLE, atp_name],
        cwd=working_directory,
    )

    process.wait()

    print("runATP.exe terminou. Aguardando geração do .lis...")

    # -------------------------
    # Esperar .lis aparecer
    # -------------------------

    while True:
        if lis_lower.exists():
            lis_path = lis_lower
            break

        if lis_upper.exists():
            lis_path = lis_upper
            break

        if time.time() - start_time > timeout:
            raise TimeoutError("Timeout: arquivo .lis não foi gerado")

        time.sleep(0.5)

    print("LIS detectado:", lis_path)

    # -------------------------
    # Esperar arquivo parar de crescer
    # -------------------------

    print("Aguardando finalização da escrita do .lis...")

    last_size = -1
    stable_checks = 0

    while stable_checks < 3:
        current_size = lis_path.stat().st_size

        if current_size == last_size:
            stable_checks += 1
        else:
            stable_checks = 0

        last_size = current_size
        time.sleep(1)

        if time.time() - start_time > timeout:
            raise TimeoutError("Timeout aguardando finalização do .lis")

    print("Arquivo .lis finalizado")

    # -------------------------
    # Verificar se simulação terminou corretamente
    # -------------------------

    print("Verificando status da simulação...")

    try:
        with lis_path.open("r", errors="ignore") as f:
            tail = f.readlines()[-50:]  # últimas linhas
    except Exception:
        tail = []

    text_tail = "".join(tail)

    if "TOTAL ELAPSED TIME" in text_tail or "Job completed" in text_tail:
        print("Simulação finalizada com sucesso")
    else:
        print("Aviso: fim da simulação não confirmado no .lis")

    end_time = time.time()
    elapsed = end_time - start_time

    print(f"Tempo total: {elapsed:.2f} s")
    print("===== ATP SIMULATION END =====\n")

    return str(lis_path)