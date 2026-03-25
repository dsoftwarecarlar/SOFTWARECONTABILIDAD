from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ProcessRequest:
    input_paths: list[Path]
    output_path: Path
    template_path: Path | None = None
    options: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ProcessResult:
    success: bool
    output_path: Path
    label: str
    metadata: dict[str, Any] = field(default_factory=dict)

