from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Optional


@dataclass
class WorkflowContext:
    requirement: str
    doc_path: str
    save_dir: str
    language: str = "zh"
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class StepResult:
    ok: bool
    step: str
    data: Dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None

