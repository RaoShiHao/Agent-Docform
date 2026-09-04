from __future__ import annotations

from typing import Any, Dict, List

from workflow.tools.dynamic_registry import DynamicToolRegistry


class ActionExecutor:
    def __init__(self, registry: DynamicToolRegistry) -> None:
        self.registry = registry

    def execute(self, actions: List[Dict[str, Any]], args_replace: Dict[str, Any]) -> List[Dict[str, Any]]:
        status_result = []
        for idx, action in enumerate(actions):
            try:
                tool_name = action.get("tool_name")
                function_name = action.get("function")
                parameters = dict(action.get("parameters") or {})
                for k, v in args_replace.items():
                    if k in parameters:
                        parameters[k] = v
                tool = self.registry.get(tool_name)
                if tool is None:
                    status_result.append(
                        {
                            "status": "error",
                            "code": "TOOL_NOT_FOUND",
                            "message": f"tool not found: {tool_name}",
                            "action_index": idx,
                            "action": action,
                        }
                    )
                    continue
                method = getattr(tool, function_name)
                status = self.registry.safe_call(method, **parameters)
                status_result.append(status)
            except Exception as e:
                status_result.append(
                    {
                        "status": "error",
                        "code": "ACTION_EXECUTE_EXCEPTION",
                        "message": str(e),
                        "action_index": idx,
                        "action": action,
                    }
                )
        return status_result

