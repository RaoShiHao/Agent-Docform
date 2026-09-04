from __future__ import annotations

import os
from typing import Any, Dict

import yaml

from constant import ABS_DIR


class PromptHub:
    def __init__(self, prompt_config_path: str) -> None:
        abs_path = os.path.join(ABS_DIR, prompt_config_path) if not os.path.isabs(prompt_config_path) else prompt_config_path
        with open(abs_path, encoding="utf-8") as f:
            self._data: Dict[str, Any] = (yaml.safe_load(f) or {}).get("prompt_config", {})

    def get(self, key: str, language: str = "zh") -> Any:
        node = self._data.get(key)
        if isinstance(node, dict):
            if language in node:
                return node.get(language)
        return node

    def resolve_prompt(self, key: str, language: str, branch: str | None = None) -> str:
        """Resolve a string prompt; supports nested keys like ``directly_analysis.table.zh``."""
        node = self._data.get(key)
        if branch and isinstance(node, dict) and branch in node:
            sub = node[branch]
            if isinstance(sub, dict):
                return str(sub.get(language) or sub.get("zh") or "").strip()
            if isinstance(sub, str):
                return sub.strip()
            return ""
        if isinstance(node, dict) and language in node:
            val = node[language]
            return val.strip() if isinstance(val, str) else ""
        if isinstance(node, str):
            return node.strip()
        return ""

