from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Any, Dict

import yaml

from constant import ABS_DIR

_ENV_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")


def _expand_env(value: Any) -> Any:
    """Recursively expand ``${ENV_VAR}`` placeholders from the process environment."""
    if isinstance(value, dict):
        return {k: _expand_env(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_expand_env(v) for v in value]
    if isinstance(value, str):
        def repl(match: re.Match[str]) -> str:
            return os.environ.get(match.group(1), match.group(0))

        return _ENV_PATTERN.sub(repl, value)
    return value


@dataclass
class DomainConfig:
    raw: Dict[str, Any]
    path: str

    @property
    def llm_config(self) -> Dict[str, Any]:
        return self.raw.get("llm_config", {})

    @property
    def vllm_config(self) -> Dict[str, Any]:
        return self.raw.get("vllm_config", {})

    @property
    def tools_config(self) -> Dict[str, Any]:
        return self.raw.get("tools_config", {})

    @property
    def refine_config(self) -> Dict[str, Any]:
        return self.raw.get("refine_config", {})

    @property
    def understand_config(self) -> Dict[str, Any]:
        return self.raw.get("understand_config", {})

    @property
    def prompt_config_path(self) -> str:
        return self.raw.get("prompt_config", "")


def load_domain_config(config_path: str) -> DomainConfig:
    abs_path = os.path.join(ABS_DIR, config_path) if not os.path.isabs(config_path) else config_path
    with open(abs_path, encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}
    return DomainConfig(raw=_expand_env(raw), path=abs_path)
