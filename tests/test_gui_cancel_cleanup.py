from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from lis_analysis.gui import ModernLisAnalysisApp


def test_single_cancel_removes_partial_outputs_and_preserves_source(tmp_path: Path):
    source_atp = tmp_path / "caso.atp"
    execution_atp = tmp_path / "caso_param.atp"
    generated_lis = tmp_path / "caso_param.lis"
    simulation_dir = tmp_path / "resultados" / "2026-07-15_12-39-18"
    simulation_dir.mkdir(parents=True)

    source_atp.write_text("ATP ORIGINAL", encoding="utf-8")
    execution_atp.write_text("ATP PARAMETRIZADO", encoding="utf-8")
    generated_lis.write_text("LIS PARCIAL", encoding="utf-8")
    (simulation_dir / "grafico_parcial.png").write_bytes(b"partial")

    ModernLisAnalysisApp._cleanup_cancelled_atp_artifacts(
        None,
        source_atp=source_atp,
        execution_atp=execution_atp,
        generated_lis=generated_lis,
        simulation_dir=simulation_dir,
    )

    assert source_atp.read_text(encoding="utf-8") == "ATP ORIGINAL"
    assert not execution_atp.exists()
    assert not generated_lis.exists()
    assert not simulation_dir.exists()
    assert not list((tmp_path / "resultados").glob("CANCELADA_*"))

def test_failure_risk_report_uses_accented_utf8_text(tmp_path: Path):
    report = ModernLisAnalysisApp._write_failure_risk_report(
        tmp_path,
        "Simulação ATP",
        [
            {
                "label": "caso.lis",
                "risk": 4.3021139614e-7,
                "corrected_cfo_kv": 1046.4206,
                "z_score": 4.783715,
                "mean_overvoltage_kv": 708.9225,
                "switching_std_kv": 67.6520,
            }
        ],
    )

    assert report is not None
    raw = report.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")

    text = raw.decode("utf-8-sig")
    assert "RELATÓRIO DE RISCO DE FALHA" in text
    assert "Contexto: Simulação ATP" in text
    assert "Índice normalizado Z" in text
    assert "Sobretensão média" in text
    assert "Desvio-padrão da sobretensão" in text
