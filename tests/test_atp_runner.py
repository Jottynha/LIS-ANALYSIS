from pathlib import Path

from lis_analysis.solver.atp_runner import (
    ATP_DIRECT_SUPPORT_FILES,
    _cleanup_direct_solver_support,
    _discover_atp_executables,
    _is_atp_temporary_file,
    _is_cancelled_atp_result,
    _lis_has_completion_marker,
    _remove_files_changed_since_snapshot,
    _snapshot_matching_files,
    _stage_direct_solver_support,
)


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


def test_direct_solver_support_preserves_existing_files(tmp_path: Path):
    support_dir = tmp_path / "support"
    working_dir = tmp_path / "work"
    support_dir.mkdir()
    working_dir.mkdir()
    for filename in ATP_DIRECT_SUPPORT_FILES:
        (support_dir / filename).write_text(f"source-{filename}", encoding="utf-8")

    existing = working_dir / ATP_DIRECT_SUPPORT_FILES[0]
    existing.write_text("user-file", encoding="utf-8")

    copied = _stage_direct_solver_support(working_dir, support_dir)

    assert existing.read_text(encoding="utf-8") == "user-file"
    assert existing not in copied
    assert len(copied) == len(ATP_DIRECT_SUPPORT_FILES) - 1

    _cleanup_direct_solver_support(copied)

    assert existing.exists()
    assert all(not path.exists() for path in copied)


def test_direct_solver_support_rolls_back_partial_copy(tmp_path: Path):
    support_dir = tmp_path / "support"
    working_dir = tmp_path / "work"
    support_dir.mkdir()
    working_dir.mkdir()
    first = ATP_DIRECT_SUPPORT_FILES[0]
    (support_dir / first).write_text("available", encoding="utf-8")

    try:
        _stage_direct_solver_support(working_dir, support_dir)
    except FileNotFoundError:
        pass
    else:
        raise AssertionError("missing support file should fail")

    assert not (working_dir / first).exists()


def test_discovers_executables_in_nonstandard_atp_root(tmp_path: Path):
    root = tmp_path / "custom-atp"
    wrapper = root / "tools" / "runATP.exe"
    direct = root / "atpmingw" / "tpbig.exe"
    wrapper.parent.mkdir(parents=True)
    direct.parent.mkdir(parents=True)
    wrapper.write_bytes(b"wrapper")
    direct.write_bytes(b"solver")

    found_wrapper, found_direct = _discover_atp_executables([root])

    assert found_wrapper == wrapper
    assert found_direct == direct


def test_temporary_cleanup_removes_only_new_or_changed_atp_scratch_files(tmp_path: Path):
    unchanged = tmp_path / "old.tmp"
    changed = tmp_path / "dum9.bin"
    pl4_result = tmp_path / "caso.pl4"
    unchanged.write_text("keep", encoding="utf-8")
    changed.write_text("old", encoding="utf-8")
    pl4_result.write_text("result", encoding="utf-8")

    snapshot = _snapshot_matching_files(tmp_path, _is_atp_temporary_file)
    changed.write_text("new content", encoding="utf-8")
    created = tmp_path / "123456.tmp"
    created.write_text("scratch", encoding="utf-8")

    removed = _remove_files_changed_since_snapshot(
        tmp_path,
        snapshot,
        _is_atp_temporary_file,
    )

    assert set(removed) == {changed, created}
    assert unchanged.exists()
    assert pl4_result.exists()


def test_cancelled_result_cleanup_preserves_unchanged_previous_results(tmp_path: Path):
    unchanged = tmp_path / "anterior.lis"
    changed = tmp_path / "caso.pl4"
    debug = tmp_path / "caso.dbg"
    unchanged.write_text("valid old result", encoding="utf-8")
    changed.write_text("old plot", encoding="utf-8")
    debug.write_text("old debug", encoding="utf-8")

    snapshot = _snapshot_matching_files(tmp_path, _is_cancelled_atp_result)
    changed.write_text("incomplete plot from cancelled run", encoding="utf-8")
    partial = tmp_path / "parcial.lis"
    partial.write_text("incomplete", encoding="utf-8")

    removed = _remove_files_changed_since_snapshot(
        tmp_path,
        snapshot,
        _is_cancelled_atp_result,
    )

    assert set(removed) == {changed, partial}
    assert unchanged.exists()
    assert debug.exists()
