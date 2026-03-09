"""Execucao de simulacoes ATP a partir de arquivos .atp."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class ATPSimulationResult:
    status: str
    lis_path: Optional[Path]
    log_path: Optional[Path]
    returncode: Optional[int]
    stdout: str
    stderr: str
    applied_params: Dict[str, int]
    warnings: List[str]


def run_atp_simulation(
    atp_path: Path,
    solver_path: str,
    params: Optional[Dict[str, Any]] = None,
    output_dir: Optional[Path] = None,
    timeout_sec: int = 300,
) -> ATPSimulationResult:
    atp_path = Path(atp_path)
    out_dir = Path(output_dir) if output_dir else atp_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    logs_dir = out_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    warnings: List[str] = []
    applied_params: Dict[str, int] = {}

    if not atp_path.exists():
        return _result_error(
            "atp_not_found",
            None,
            None,
            None,
            "",
            f"Arquivo nao encontrado: {atp_path}",
            applied_params,
            warnings,
        )

    solver = _resolve_solver(solver_path)
    if not solver:
        return _result_error(
            "solver_not_found",
            None,
            None,
            None,
            "",
            f"Executavel ATP nao encontrado: {solver_path}",
            applied_params,
            warnings,
        )

    try:
        atp_text = atp_path.read_text(encoding="windows-1252", errors="ignore")
    except Exception:
        atp_text = atp_path.read_text(encoding="latin-1", errors="ignore")

    if params:
        atp_text, applied_params, warnings = _apply_params(atp_text, params)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stage_dir = Path(tempfile.mkdtemp(prefix="atp_stage_"))
    try:
        deck_name = _safe_deck_name(atp_path.stem) + atp_path.suffix
        deck_path = stage_dir / deck_name

        deck_text = _normalize_line_endings(atp_text)
        deck_path.write_text(deck_text, encoding="windows-1252", errors="ignore")

        is_wrapper = _is_runatp_wrapper(solver)
        deck_arg = str(deck_path) if is_wrapper else deck_path.name
        run_cwd = Path(solver).parent if is_wrapper else stage_dir
        search_dirs = [stage_dir, run_cwd, atp_path.parent, out_dir]
        before_files = _snapshot_dirs(search_dirs)

        cmd = _build_command(solver, deck_arg)
        log_path = logs_dir / f"{atp_path.stem}_{timestamp}.log"
        _write_log(
            log_path,
            "running",
            cmd,
            run_cwd,
            None,
            "",
            "",
            None,
            applied_params,
            warnings,
        )
        stdout, stderr, returncode = _run_command(cmd, run_cwd, timeout_sec)

        lis_path = _pick_lis(search_dirs, before_files, [deck_path.stem, atp_path.stem])
        moved_lis = None
        if lis_path and lis_path.exists() and lis_path.stat().st_size > 0:
            moved_lis = _move_preserve_name(lis_path, out_dir, timestamp)

        status = "success" if moved_lis else "no_lis"
        if returncode not in (0, None):
            status = "error_with_lis" if moved_lis else "error"

        _write_log(
            log_path,
            status,
            cmd,
            run_cwd,
            moved_lis,
            stdout,
            stderr,
            returncode,
            applied_params,
            warnings,
        )

        return ATPSimulationResult(
            status=status,
            lis_path=moved_lis,
            log_path=log_path,
            returncode=returncode,
            stdout=stdout,
            stderr=stderr,
            applied_params=applied_params,
            warnings=warnings,
        )
    finally:
        for _ in range(3):
            try:
                shutil.rmtree(stage_dir, ignore_errors=False)
                break
            except Exception:
                time.sleep(0.5)
        else:
            warnings.append(f"Staging nao removido: {stage_dir}")


def _apply_params(atp_text: str, params: Dict[str, Any]) -> Tuple[str, Dict[str, int], List[str]]:
    applied: Dict[str, int] = {}
    warnings: List[str] = []
    text = atp_text

    for key, value in params.items():
        if value is None:
            warnings.append(f"Parametro '{key}' ignorado (valor nulo).")
            continue
        if isinstance(value, (dict, list)):
            warnings.append(f"Parametro '{key}' ignorado (tipo nao suportado).")
            continue

        count = 0
        for pattern in _param_patterns(str(key)):
            text, replaced = _replace_param(text, pattern, value)
            count += replaced

        applied[str(key)] = count
        if count == 0:
            warnings.append(f"Parametro '{key}' nao encontrado no deck.")

    return text, applied, warnings


def _param_patterns(key: str) -> List[re.Pattern]:
    esc = re.escape(key)
    return [
        re.compile(rf"(?i)(\\b{esc}\\b\\s*=\\s*)([-+]?\\d*\\.?\\d+(?:[eE][-+]?\\d+)?)"),
        re.compile(rf"(?i)(\\b{esc}\\b\\s+)([-+]?\\d*\\.?\\d+(?:[eE][-+]?\\d+)?)"),
    ]


def _replace_param(text: str, pattern: re.Pattern, value: Any) -> Tuple[str, int]:
    replaced = 0

    def repl(match: re.Match) -> str:
        nonlocal replaced
        old = match.group(2)
        new = _format_value(value, len(old))
        replaced += 1
        return f"{match.group(1)}{new}"

    new_text = pattern.sub(repl, text)
    return new_text, replaced


def _format_value(value: Any, width: int) -> str:
    if isinstance(value, str):
        formatted = value
    else:
        try:
            formatted = f"{float(value):g}"
        except Exception:
            formatted = str(value)
    if width > 0 and len(formatted) < width:
        return formatted.rjust(width)
    return formatted


def _normalize_line_endings(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text.replace("\n", "\r\n")


def _safe_deck_name(stem: str) -> str:
    return re.sub(r"[=\\s]+", "_", stem)


def _build_command(solver: str, deck_name: str) -> List[str]:
    ext = Path(solver).suffix.lower()
    if ext in [".bat", ".cmd"]:
        if os.name == "nt":
            return ["cmd", "/c", solver, deck_name]
        if shutil.which("wine"):
            return ["wine", "cmd", "/c", solver, deck_name]
        raise RuntimeError("Wine nao encontrado para executar .bat/.cmd")
    return [solver, deck_name]


def _is_runatp_wrapper(solver: str) -> bool:
    return "runatp" in Path(solver).name.lower()


def _run_command(cmd: List[str], cwd: Path, timeout_sec: int) -> Tuple[str, str, Optional[int]]:
    stdout = ""
    stderr = ""
    returncode = None
    proc = None
    try:
        if os.name == "nt":
            proc = subprocess.Popen(
                cmd,
                cwd=cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                stdin=subprocess.PIPE,
                text=True,
            )
        else:
            import os as _os

            proc = subprocess.Popen(
                cmd,
                cwd=cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                stdin=subprocess.PIPE,
                text=True,
                preexec_fn=_os.setsid,
            )
        stdout, stderr = proc.communicate(input=("go\n" * 3), timeout=timeout_sec)
        returncode = proc.returncode
    except subprocess.TimeoutExpired:
        if proc is not None:
            try:
                if os.name == "nt":
                    subprocess.run(
                        ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                    )
                else:
                    import os as _os
                    import signal as _signal

                    _os.killpg(_os.getpgid(proc.pid), _signal.SIGKILL)
            except Exception:
                pass
        returncode = -9
        stderr = (stderr or "") + "\n[timeout] Processo excedeu o tempo limite."
    except Exception as e:
        returncode = -1
        stderr = f"Falha ao executar ATP: {e}"
    return stdout, stderr, returncode


def _snapshot_dirs(dirs: List[Path]) -> Dict[str, set]:
    snapshot: Dict[str, set] = {}
    for d in dirs:
        if d.exists() and d.is_dir():
            try:
                snapshot[str(d)] = set(p.name for p in d.iterdir() if p.is_file())
            except Exception:
                snapshot[str(d)] = set()
    return snapshot


def _pick_lis(search_dirs: List[Path], before_files: Dict[str, set], expected_stems: List[str]) -> Optional[Path]:
    candidates: List[Path] = []
    for d in search_dirs:
        if d.exists() and d.is_dir():
            candidates.extend(list(d.glob("*.lis")))
            candidates.extend(list(d.glob("*.LIS")))
    if not candidates:
        return None

    # Preferir arquivos novos e com stem esperado
    expected = set(s.lower() for s in expected_stems)
    scored: List[Tuple[int, float, Path]] = []
    for p in candidates:
        try:
            is_new = 1 if p.name not in before_files.get(str(p.parent), set()) else 0
            stem_match = 1 if p.stem.lower() in expected else 0
            score = (is_new * 2) + stem_match
            scored.append((score, p.stat().st_mtime, p))
        except Exception:
            continue

    if not scored:
        return None
    scored.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return scored[0][2]


def _move_preserve_name(src: Path, dst_dir: Path, ts: str) -> Path:
    target = dst_dir / src.name
    if target.exists():
        target = dst_dir / f"{src.stem}_{ts}{src.suffix}"
    return Path(shutil.move(str(src), str(target)))


def _write_log(
    log_path: Path,
    status: str,
    cmd: List[str],
    cwd: Path,
    lis_path: Optional[Path],
    stdout: str,
    stderr: str,
    returncode: Optional[int],
    applied_params: Dict[str, int],
    warnings: List[str],
) -> None:
    lines = [
        f"Status: {status}",
        f"Return code: {returncode}",
        f"CWD: {cwd}",
        f"Command: {' '.join(cmd)}",
        f"LIS: {lis_path if lis_path else '(none)'}",
    ]
    if applied_params:
        lines.append(f"Applied params: {applied_params}")
    if warnings:
        lines.append("Warnings:")
        lines.extend([f"  - {w}" for w in warnings])
    lines.extend([
        "---- STDOUT ----",
        stdout or "(vazio)",
        "---- STDERR ----",
        stderr or "(vazio)",
    ])
    try:
        log_path.write_text("\n".join(lines), encoding="utf-8")
    except Exception:
        pass


def _resolve_solver(solver_path: str) -> Optional[str]:
    if not solver_path:
        return None
    if Path(solver_path).exists():
        return str(Path(solver_path))
    return shutil.which(solver_path)


def _result_error(
    status: str,
    lis_path: Optional[Path],
    log_path: Optional[Path],
    returncode: Optional[int],
    stdout: str,
    stderr: str,
    applied_params: Dict[str, int],
    warnings: List[str],
) -> ATPSimulationResult:
    return ATPSimulationResult(
        status=status,
        lis_path=lis_path,
        log_path=log_path,
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
        applied_params=applied_params,
        warnings=warnings,
    )
