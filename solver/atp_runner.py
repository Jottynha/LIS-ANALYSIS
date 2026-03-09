"""Runner simples para executar o solver ATP via runATP.exe."""

from __future__ import annotations

import subprocess
import time
from pathlib import Path

ATP_EXECUTABLE = r"C:\ATP\tools\runATP.exe"


def run_atp_solver(atp_file_path: str) -> str:
    """
    Executes ATP using runATP.exe and returns the path to the generated .lis file.
    """
    atp_path = Path(atp_file_path)

    if not atp_path.exists() or not atp_path.is_file():
        raise FileNotFoundError(f"Arquivo .atp nao encontrado: {atp_file_path}")

    working_directory = atp_path.parent
    atp_file_name = atp_path.name
    base_name = atp_path.stem
    lis_file_path = working_directory / f"{base_name}.lis"
    lis_file_path_upper = working_directory / f"{base_name}.LIS"

    print("Running ATP simulation")
    print("ATP executable:", ATP_EXECUTABLE)
    print("ATP input:", str(atp_path))
    print("Working directory:", str(working_directory))

    start_time = time.time()
    process = subprocess.Popen(
        [ATP_EXECUTABLE, atp_file_name],
        cwd=working_directory,
    )
    return_code = process.wait()
    end_time = time.time()

    execution_time = end_time - start_time
    print("ATP finished")
    print("Execution time:", execution_time)

    if lis_file_path.exists():
        return str(lis_file_path)

    if lis_file_path_upper.exists():
        return str(lis_file_path_upper)

    if return_code != 0:
        raise RuntimeError(
            f"ATP execution failed with return code {return_code} and .lis file was not generated"
        )

    raise RuntimeError("ATP execution finished but the .lis file was not generated")
