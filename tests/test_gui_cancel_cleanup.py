from pathlib import Path

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
