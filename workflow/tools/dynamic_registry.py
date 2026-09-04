from __future__ import annotations

import importlib.util
import inspect
import os
import sys
from pathlib import Path
from typing import Any, Dict

from constant import ABS_DIR


class DynamicToolRegistry:
    def __init__(self, tools_config: Dict[str, Dict[str, str]]) -> None:
        self.tools: Dict[str, Any] = {}
        for _, info in (tools_config or {}).items():
            rel_path = info.get("tool_path", "")
            class_name = info.get("tool_name", "")
            if rel_path and class_name:
                self.tools[class_name] = self._import_tool(rel_path, class_name)

    def _import_tool(self, relative_path: str, class_name: str) -> Any:
        abs_path = os.path.normpath(os.path.join(ABS_DIR, relative_path.lstrip("/")))
        module_path = Path(abs_path).with_suffix(".py")
        spec = importlib.util.spec_from_file_location(module_path.stem, module_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Cannot load tool module: {module_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_path.stem] = module
        spec.loader.exec_module(module)
        if not hasattr(module, class_name):
            raise AttributeError(f"{class_name} not found in {module_path}")
        return getattr(module, class_name)()

    def get(self, tool_name: str) -> Any:
        return self.tools.get(tool_name)

    def describe(self, language: str = "zh", tool_name: str | None = None) -> str:
        if language not in ("zh", "en"):
            language = "zh"
        if tool_name:
            tools = {tool_name: self.tools.get(tool_name)} if tool_name in self.tools else {}
        else:
            tools = self.tools
        blocks = []
        for name, tool in tools.items():
            if tool is None:
                continue
            desc = tool.config.get("description", {}).get(language, "")
            methods = tool.describe_tools(language)
            blocks.append(f"- {name}\n  - description: {desc}\n  - methods: {methods}")
        return "\n".join(blocks)

    @staticmethod
    def safe_call(func, **kwargs):
        sig = inspect.signature(func)
        valid = {k: v for k, v in kwargs.items() if k in sig.parameters}
        return func(**valid)

