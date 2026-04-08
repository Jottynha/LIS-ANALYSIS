from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ATPElement:
    """Representa um componente editavel de um arquivo ATP."""

    type: str
    name: str
    nodes: list[str]
    parameters: dict[str, Any]
    raw_line: str
    line_index: int
    line_end_index: int
    continuation_lines: list[str] = field(default_factory=list)
    modified: bool = False

    def parameter_name(self) -> str:
        return "value"

    def get_value(self) -> float:
        key = self.parameter_name()
        value = self.parameters.get(key)
        if value is None:
            raise ValueError(f"Parametro '{key}' ausente em {self.name}")
        return float(value)

    def set_value(self, new_value: float) -> None:
        key = self.parameter_name()
        self.parameters[key] = float(new_value)
        self.modified = True


@dataclass
class Resistor(ATPElement):
    def parameter_name(self) -> str:
        return "resistance"


@dataclass
class Inductor(ATPElement):
    def parameter_name(self) -> str:
        return "inductance"


@dataclass
class Capacitor(ATPElement):
    def parameter_name(self) -> str:
        return "capacitance"


@dataclass
class VoltageSource(ATPElement):
    def parameter_name(self) -> str:
        return "voltage"


@dataclass
class CurrentSource(ATPElement):
    def parameter_name(self) -> str:
        return "current"


TYPE_TO_CLASS = {
    "R": Resistor,
    "L": Inductor,
    "C": Capacitor,
    "V": VoltageSource,
    "I": CurrentSource,
}


def element_class_for(card_type: str):
    return TYPE_TO_CLASS.get(card_type.upper(), ATPElement)
