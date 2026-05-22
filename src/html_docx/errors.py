from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class HDocxError(Exception):
    code: str
    message: str
    details: dict[str, Any] = field(default_factory=dict)

    def __str__(self) -> str:
        return f"{self.code}: {self.message}"

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "code": self.code,
            "severity": "error",
            "message": self.message,
        }
        if self.details:
            payload["details"] = self.details
        return payload
