from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Set, Tuple

from workflow.tools.dynamic_registry import DynamicToolRegistry


@dataclass
class ValidationError:
    code: str
    message: str
    action_index: int | None = None
    action: Dict[str, Any] | None = None


class ActionValidator:
    """LLM outputs sometimes follow prompt typos (e.g. ``PageTool`` vs registered ``PageTools``)."""

    _TOOL_NAME_ALIASES: Dict[str, str] = {
        "PageTool": "PageTools",
        "TextTool": "TextTools",
        "TableTool": "TableTools",
        "ImageTool": "ImageTools",
    }

    def __init__(self, registry: DynamicToolRegistry, allowed_tools: Set[str]) -> None:
        self.registry = registry
        self.allowed_tools = allowed_tools

    def validate_actions(self, actions: Any) -> Tuple[List[Dict[str, Any]], List[ValidationError]]:
        errors: List[ValidationError] = []
        if not isinstance(actions, list):
            errors.append(
                ValidationError(
                    code="INVALID_ACTIONS_TYPE",
                    message="funcall output must be a JSON array of actions",
                )
            )
            return [], errors

        valid_actions: List[Dict[str, Any]] = []
        for idx, action in enumerate(actions):
            if not isinstance(action, dict):
                errors.append(
                    ValidationError(
                        code="INVALID_ACTION_TYPE",
                        message="each action must be a JSON object",
                        action_index=idx,
                        action={"raw": action},
                    )
                )
                continue

            tool_name = action.get("tool_name")
            function_name = action.get("function")
            parameters = action.get("parameters")

            if not tool_name or not isinstance(tool_name, str):
                errors.append(
                    ValidationError(
                        code="MISSING_TOOL_NAME",
                        message="action.tool_name is required and must be string",
                        action_index=idx,
                        action=action,
                    )
                )
                continue
            canonical = self._TOOL_NAME_ALIASES.get(tool_name.strip())
            if canonical and canonical in self.allowed_tools:
                action["tool_name"] = canonical
                tool_name = canonical
            if tool_name not in self.allowed_tools:
                errors.append(
                    ValidationError(
                        code="TOOL_NOT_ALLOWED",
                        message=f"tool '{tool_name}' is not allowed for this domain",
                        action_index=idx,
                        action=action,
                    )
                )
                continue

            tool = self.registry.get(tool_name)
            if tool is None:
                errors.append(
                    ValidationError(
                        code="TOOL_NOT_FOUND",
                        message=f"tool '{tool_name}' is not registered",
                        action_index=idx,
                        action=action,
                    )
                )
                continue

            if not function_name or not isinstance(function_name, str):
                errors.append(
                    ValidationError(
                        code="MISSING_FUNCTION",
                        message="action.function is required and must be string",
                        action_index=idx,
                        action=action,
                    )
                )
                continue
            if not hasattr(tool, function_name):
                errors.append(
                    ValidationError(
                        code="FUNCTION_NOT_FOUND",
                        message=f"function '{function_name}' not found on tool '{tool_name}'",
                        action_index=idx,
                        action=action,
                    )
                )
                continue

            if parameters is None:
                action["parameters"] = {}
                parameters = action["parameters"]
            if not isinstance(parameters, dict):
                errors.append(
                    ValidationError(
                        code="INVALID_PARAMETERS",
                        message="action.parameters must be an object",
                        action_index=idx,
                        action=action,
                    )
                )
                continue

            valid_actions.append(action)

        return valid_actions, errors

