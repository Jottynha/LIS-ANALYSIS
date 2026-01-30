"""Execução ATP a partir de arquivos .acp."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, List


@dataclass
class ATPResult:
    status: str
    lis_path: Optional[Path]
    dbg_path: Optional[Path]
    log_path: Optional[Path]
    returncode: Optional[int]
    stdout: str
    stderr: str


class ATPRunner:
    def __init__(self, solver_path: str, timeout_sec: int = 300):
        self.solver_path = solver_path
        self.timeout_sec = timeout_sec

    def run_acp(self, acp_path: Path, output_dir: Path) -> ATPResult:
        acp_path = Path(acp_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        logs_dir = output_dir / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)

        if not acp_path.exists():
            return self._result_error("acp_not_found", None, None, None, None, "", f"Arquivo não encontrado: {acp_path}")

        solver = self._resolve_solver(self.solver_path)
        if not solver:
            return self._result_error("solver_not_found", None, None, None, None, "", f"Executável ATP não encontrado: {self.solver_path}")

        try:
            atp_text, atp_bytes, atp_entry, atp_score = self._extract_atp_from_acp(acp_path)
        except Exception as e:
            return self._result_error("extract_failed", None, None, None, None, "", f"Falha ao extrair ATP: {e}")

        if atp_score < 0.6:
            return self._result_error(
                "acp_binary",
                None,
                None,
                None,
                None,
                "",
                "Conteúdo .$$$ parece binário. Exporte o .atp pelo ATPDraw e use a opção .atp."
            )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with tempfile.TemporaryDirectory(prefix="atp_stage_") as tmpdir:
            stage_dir = Path(tmpdir)
            deck_name = self._safe_deck_name(acp_path.stem) + ".atp"
            deck_path = stage_dir / deck_name
            use_text = self._should_reencode(atp_bytes)
            if use_text:
                deck_text = atp_text.replace("\t", " ")
                deck_text = deck_text.replace("\r\n", "\n").replace("\r", "\n")
                deck_text = deck_text.replace("\n", "\r\n")
                deck_path.write_text(deck_text, encoding="windows-1252", errors="ignore")
                self._write_startup_override(stage_dir)
            else:
                deck_bytes = atp_bytes.replace(b"\x00", b"")
                deck_bytes = deck_bytes.replace(b"\t", b" ")
                deck_path.write_bytes(deck_bytes)

            self._copy_includes(atp_text, acp_path.parent, stage_dir)
            self._copy_startup(Path(solver).parent, stage_dir)

            cmd = self._build_command(solver, deck_path.name)
            stdout, stderr, returncode = self._run_command(cmd, stage_dir)

            lis_path = self._pick_lis(stage_dir)
            moved_lis = None
            moved_dbg = None

            if lis_path and lis_path.exists() and lis_path.stat().st_size > 0:
                moved_lis = self._move_preserve_name(lis_path, output_dir, timestamp)

            for dbg in stage_dir.glob("*.dbg"):
                moved_dbg = self._move_preserve_name(dbg, output_dir, timestamp)
                break

            status = "success" if moved_lis else "no_lis"
            if returncode not in (0, None):
                status = "error_with_lis" if moved_lis else "error"

            log_path = logs_dir / f"{acp_path.stem}_{timestamp}.log"
            self._write_log(
                log_path,
                status,
                cmd,
                stage_dir,
                moved_lis,
                moved_dbg,
                stdout,
                stderr,
                returncode,
                extra_lines=[f"ACP entry: {atp_entry}"]
            )

            return ATPResult(status, moved_lis, moved_dbg, log_path, returncode, stdout, stderr)

    def _extract_atp_from_acp(self, acp_path: Path) -> tuple[str, bytes, str, float]:

        with zipfile.ZipFile(acp_path, "r") as zip_ref:
            candidates = [f for f in zip_ref.namelist() if f.lower().endswith(".$$$")]
            if not candidates:
                raise ValueError("Arquivo .$$$ não encontrado dentro do .acp")

            scored = []
            for name in candidates:
                data = zip_ref.read(name)
                scored.append((self._score_bytes(data), name, data))
            scored.sort(key=lambda t: t[0], reverse=True)
            best_score, best_name, content = scored[0]
        text = None
        # Detectar UTF-16 (muitos NULs alternados)
        if content and content.count(b"\x00") > len(content) * 0.2:
            try:
                text = content.decode("utf-16le")
            except Exception:
                try:
                    text = content.decode("utf-16be")
                except Exception:
                    text = None
        if text is None:
            try:
                text = content.decode("windows-1252")
            except Exception:
                text = content.decode("latin-1", errors="ignore")
        return text, content, best_name, best_score

    def run_atp(self, atp_path: Path, output_dir: Path) -> ATPResult:
        atp_path = Path(atp_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        logs_dir = output_dir / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)

        if not atp_path.exists():
            return self._result_error("atp_not_found", None, None, None, None, "", f"Arquivo não encontrado: {atp_path}")

        solver = self._resolve_solver(self.solver_path)
        if not solver:
            return self._result_error("solver_not_found", None, None, None, None, "", f"Executável ATP não encontrado: {self.solver_path}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with tempfile.TemporaryDirectory(prefix="atp_stage_") as tmpdir:
            stage_dir = Path(tmpdir)
            deck_name = self._safe_deck_name(atp_path.stem) + atp_path.suffix
            deck_path = stage_dir / deck_name
            shutil.copy2(atp_path, deck_path)

            self._copy_startup(Path(solver).parent, stage_dir)

            cmd = self._build_command(solver, deck_path.name)
            stdout, stderr, returncode = self._run_command(cmd, stage_dir)

            lis_path = self._pick_lis(stage_dir)
            moved_lis = None
            moved_dbg = None

            if lis_path and lis_path.exists() and lis_path.stat().st_size > 0:
                moved_lis = self._move_preserve_name(lis_path, output_dir, timestamp)

            for dbg in stage_dir.glob("*.dbg"):
                moved_dbg = self._move_preserve_name(dbg, output_dir, timestamp)
                break

            status = "success" if moved_lis else "no_lis"
            if returncode not in (0, None):
                status = "error_with_lis" if moved_lis else "error"

            log_path = logs_dir / f"{atp_path.stem}_{timestamp}.log"
            self._write_log(log_path, status, cmd, stage_dir, moved_lis, moved_dbg, stdout, stderr, returncode)

            return ATPResult(status, moved_lis, moved_dbg, log_path, returncode, stdout, stderr)

    def _score_bytes(self, raw: bytes) -> float:
        if not raw:
            return 0.0
        printable = 0
        linebreaks = 0
        nul_count = 0
        for b in raw:
            if b == 0:
                nul_count += 1
            if b == 10:
                linebreaks += 1
            if b in (9, 10, 13) or 32 <= b <= 126 or b >= 160:
                printable += 1
        ratio = printable / len(raw)
        lb_score = min(linebreaks / 500.0, 1.0)
        nul_penalty = min(nul_count / len(raw), 1.0)
        try:
            text = raw.decode("windows-1252", errors="ignore").upper()
        except Exception:
            text = ""
        keyword_bonus = 0.0
        if "BEGIN NEW DATA" in text or "BEGIN" in text:
            keyword_bonus += 0.2
        if "C " in text or "$" in text:
            keyword_bonus += 0.05
        return (ratio * 0.6) + (lb_score * 0.3) + keyword_bonus - (nul_penalty * 0.6)

    def _safe_deck_name(self, stem: str) -> str:
        return re.sub(r"[=\s]+", "_", stem)

    def _copy_includes(self, atp_text: str, base_dir: Path, stage_dir: Path) -> None:
        include_pat = re.compile(r"\b(INCLUDE|\$INCLUDE|\.INC)\b", re.IGNORECASE)
        for line in atp_text.splitlines():
            if not include_pat.search(line):
                continue
            m = re.search(r"\"([^\"]+)\"|'([^']+)'", line)
            candidate = None
            if m:
                candidate = m.group(1) or m.group(2)
            else:
                parts = line.strip().split()
                candidate = parts[-1] if parts else None
            if not candidate:
                continue
            candidate_norm = candidate.replace("\\\\", os.sep)
            inc_path = Path(candidate_norm)
            if not inc_path.is_absolute():
                inc_path = (base_dir / inc_path).resolve()
            if inc_path.exists() and inc_path.is_file():
                rel_target = Path(candidate_norm)
                target = stage_dir / rel_target
                target.parent.mkdir(parents=True, exist_ok=True)
                try:
                    shutil.copy2(inc_path, target)
                except Exception:
                    pass

    def _copy_startup(self, solver_dir: Path, stage_dir: Path) -> None:
        for name in ("startup", "STARTUP"):
            candidate = solver_dir / name
            if candidate.exists() and candidate.is_file():
                try:
                    content = candidate.read_text(encoding="utf-8", errors="ignore")
                    if re.search(r"\bNOTAB\s*=\s*[1-9]", content) or re.search(r"\bUNIXON\s*=\s*[1-9]", content):
                        return
                except Exception:
                    pass
                try:
                    shutil.copy2(candidate, stage_dir / name)
                except Exception:
                    pass

    def _write_startup_override(self, stage_dir: Path) -> None:
        startup_path = stage_dir / "startup"
        try:
            startup_path.write_text("NOTAB = 0\nUNIXON = 0\n", encoding="utf-8")
        except Exception:
            pass

    def _should_reencode(self, raw: bytes) -> bool:
        if not raw:
            return True
        nul_ratio = raw.count(b"\x00") / max(1, len(raw))
        if nul_ratio > 0.01:
            return True
        printable = 0
        for b in raw:
            if b in (9, 10, 13) or 32 <= b <= 126 or b >= 160:
                printable += 1
        printable_ratio = printable / len(raw)
        return printable_ratio < 0.8

    def _build_command(self, solver: str, deck_name: str) -> List[str]:
        ext = Path(solver).suffix.lower()
        if ext in [".bat", ".cmd"]:
            if os.name == "nt":
                return ["cmd", "/c", solver, deck_name]
            if shutil.which("wine"):
                return ["wine", "cmd", "/c", solver, deck_name]
            raise RuntimeError("Wine não encontrado para executar .bat/.cmd")
        return [solver, deck_name]

    def _run_command(self, cmd: List[str], cwd: Path) -> tuple[str, str, Optional[int]]:
        stdout = ""
        stderr = ""
        returncode = None
        try:
            if os.name == "nt":
                proc = subprocess.Popen(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE, text=True)
            else:
                import os as _os
                proc = subprocess.Popen(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE, text=True, preexec_fn=_os.setsid)
            stdout, stderr = proc.communicate(input=("go\n" * 3), timeout=self.timeout_sec)
            returncode = proc.returncode
        except subprocess.TimeoutExpired:
            returncode = -9
            stderr = (stderr or "") + "\n[timeout] Processo excedeu o tempo limite."
        except Exception as e:
            returncode = -1
            stderr = f"Falha ao executar ATP: {e}"
        return stdout, stderr, returncode

    def _pick_lis(self, stage_dir: Path) -> Optional[Path]:
        candidates = list(stage_dir.glob("*.lis")) + list(stage_dir.glob("*.LIS"))
        if not candidates:
            return None
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return candidates[0]

    def _move_preserve_name(self, src: Path, dst_dir: Path, ts: str) -> Path:
        target = dst_dir / src.name
        if target.exists():
            target = dst_dir / f"{src.stem}_{ts}{src.suffix}"
        return Path(shutil.move(str(src), str(target)))

    def _write_log(
        self,
        log_path: Path,
        status: str,
        cmd: List[str],
        cwd: Path,
        lis_path: Optional[Path],
        dbg_path: Optional[Path],
        stdout: str,
        stderr: str,
        returncode: Optional[int],
        extra_lines: Optional[List[str]] = None,
    ) -> None:
        lines = [
            f"Status: {status}",
            f"Return code: {returncode}",
            f"CWD: {cwd}",
            f"Command: {' '.join(cmd)}",
            f"LIS: {lis_path if lis_path else '(none)'}",
            f"DBG: {dbg_path if dbg_path else '(none)'}",
        ]
        if extra_lines:
            lines.extend(extra_lines)
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

    def _resolve_solver(self, solver_path: str) -> Optional[str]:
        if not solver_path:
            return None
        if Path(solver_path).exists():
            return str(Path(solver_path))
        resolved = shutil.which(solver_path)
        return resolved

    def _result_error(
        self,
        status: str,
        lis_path: Optional[Path],
        dbg_path: Optional[Path],
        log_path: Optional[Path],
        returncode: Optional[int],
        stdout: str,
        stderr: str,
    ) -> ATPResult:
        return ATPResult(status, lis_path, dbg_path, log_path, returncode, stdout, stderr)
