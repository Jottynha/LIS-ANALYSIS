from pathlib import Path

from lis_analysis.solver.atp_runner import _lis_has_completion_marker


def test_lis_completion_marker_requires_final_timing_block(tmp_path: Path):
    lis_path = tmp_path / "complete.lis"
    lis_path.write_text(
        "\n".join(
            [
                "Actual List Sizes for the preceding solution follow.",
                "Seconds for time-step loop : 1.0 1.0 1.0",
                "Seconds after DELTAT-loop  : 0.1 0.1 0.1",
                "                   Totals  : 1.1 1.1 1.1",
            ]
        ),
        encoding="utf-8",
    )

    assert _lis_has_completion_marker(lis_path)


def test_lis_completion_marker_rejects_partial_file(tmp_path: Path):
    lis_path = tmp_path / "partial.lis"
    lis_path.write_text(
        "Actual List Sizes for the preceding solution follow.\n"
        "Seconds for time-step loop : 1.0 1.0 1.0\n",
        encoding="utf-8",
    )

    assert not _lis_has_completion_marker(lis_path)


def test_lis_completion_marker_uses_only_tail(tmp_path: Path):
    lis_path = tmp_path / "large.lis"
    lis_path.write_text(
        "Actual List Sizes for the preceding solution follow.\n"
        "Seconds after DELTAT-loop : 0.1 0.1 0.1\n"
        "Totals : 1.1 1.1 1.1\n"
        + ("x" * (70 * 1024)),
        encoding="utf-8",
    )

    assert not _lis_has_completion_marker(lis_path)
