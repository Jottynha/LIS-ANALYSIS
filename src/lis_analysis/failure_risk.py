from __future__ import annotations

from dataclasses import dataclass
from math import erfc, exp, isfinite, sqrt


@dataclass(frozen=True)
class FailureRiskConfig:
    """Parametros geometricos e eletricos usados no calculo de Hileman."""

    base_voltage_kv: float = 429.0
    conductor_height_m: float = 17.98
    conductor_structure_distance_m: float = 3.50
    tower_width_m: float = 2.0
    subconductors_per_phase: float = 4.0
    insulation_distance_m: float = 3.50
    parallel_gaps: float = 804.0


@dataclass(frozen=True)
class FailureRiskResult:
    gap_factor: float
    cfo_kv: float
    withstand_std_kv: float
    corrected_cfo_kv: float
    corrected_withstand_std_kv: float
    mean_overvoltage_kv: float
    switching_std_kv: float
    z_score: float
    risk: float


def _require_finite(name: str, value: float) -> float:
    parsed = float(value)
    if not isfinite(parsed):
        raise ValueError(f"{name} deve ser um numero finito")
    return parsed


def calculate_failure_risk(
    mean_overvoltage_pu: float,
    switching_std_pu: float,
    config: FailureRiskConfig,
) -> FailureRiskResult:
    """Calcula o risco de falha fase-terra pelas equacoes (1)-(6) do artigo.

    A media e o desvio-padrao chegam do LIS em p.u. e sao convertidos para kV
    pela tensao-base antes de serem combinados com o CFO, que tambem esta em kV.
    """

    mean_pu = _require_finite("V50 em p.u.", mean_overvoltage_pu)
    switching_std_pu = _require_finite("sigma_s em p.u.", switching_std_pu)
    base_kv = _require_finite("tensao-base", config.base_voltage_kv)
    height = _require_finite("H", config.conductor_height_m)
    distance = _require_finite("D", config.conductor_structure_distance_m)
    width = _require_finite("S", config.tower_width_m)
    subconductors = _require_finite("N", config.subconductors_per_phase)
    insulation_distance = _require_finite("d", config.insulation_distance_m)
    parallel_gaps = _require_finite("n", config.parallel_gaps)

    if mean_pu < 0 or switching_std_pu < 0:
        raise ValueError("V50 e sigma_s nao podem ser negativos")
    if base_kv <= 0 or height <= 0 or distance <= 0 or insulation_distance <= 0:
        raise ValueError("tensao-base, H, D e d devem ser maiores que zero")
    if width < 0 or subconductors <= 0 or parallel_gaps <= 0:
        raise ValueError("S deve ser nao negativo; N e n devem ser maiores que zero")

    gap_factor = (
        1.25
        + 0.005 * ((height / distance) - 6.0)
        + 0.25 * (exp(-8.0 * width / distance) - 0.2)
        - 0.007 * (distance - 5.0)
        + 0.01 * (subconductors - 2.0)
    )
    cfo_kv = gap_factor * (3400.0 / (1.0 + 8.0 / insulation_distance))
    if gap_factor <= 0 or cfo_kv <= 0:
        raise ValueError("a geometria informada produz kg/CFO nao positivo")

    withstand_std_kv = 0.06 * cfo_kv
    fifth_root_n = parallel_gaps ** (1.0 / 5.0)
    corrected_cfo_kv = cfo_kv * (
        1.0
        - 4.0
        * (withstand_std_kv / cfo_kv)
        * (1.0 - 1.0 / fifth_root_n)
    )
    corrected_withstand_std_kv = withstand_std_kv / fifth_root_n

    mean_overvoltage_kv = mean_pu * base_kv
    switching_std_kv = switching_std_pu * base_kv
    combined_std_kv = sqrt(
        corrected_withstand_std_kv**2 + switching_std_kv**2
    )
    if combined_std_kv <= 0:
        raise ValueError("o desvio-padrao combinado deve ser maior que zero")

    z_score = (corrected_cfo_kv - mean_overvoltage_kv) / combined_std_kv
    # 1 - Phi(z) = 0.5 * erfc(z/sqrt(2)); a equacao (4) aplica outro 1/2.
    risk = 0.25 * erfc(z_score / sqrt(2.0))

    return FailureRiskResult(
        gap_factor=gap_factor,
        cfo_kv=cfo_kv,
        withstand_std_kv=withstand_std_kv,
        corrected_cfo_kv=corrected_cfo_kv,
        corrected_withstand_std_kv=corrected_withstand_std_kv,
        mean_overvoltage_kv=mean_overvoltage_kv,
        switching_std_kv=switching_std_kv,
        z_score=z_score,
        risk=max(0.0, min(0.5, risk)),
    )
