from __future__ import annotations

from typing import Any, Dict

from workflow.llm.runtime import LLMRuntime
from workflow.prompts.prompt_hub import PromptHub


class AnalysisStage:
    """
    Stages 1 & 2 (analysis side):

      - Requirement Classify  — ``requirement_classify``
      - Requirement Expansion — ``directly_analysis`` / ``interpretive_analysis``
    """

    def __init__(self, runtime: LLMRuntime, prompts: PromptHub, language: str) -> None:
        self.runtime = runtime
        self.prompts = prompts
        self.language = language

    @staticmethod
    def _extract_content(response: Dict[str, Any]) -> str:
        data = response.get("data") or {}
        return data.get("content", "") if response.get("success") else ""

    @staticmethod
    def _json_parse(json_str: Any):
        import ast
        import json
        import re
        import json5

        if isinstance(json_str, (dict, list)):
            return json_str
        if not isinstance(json_str, str):
            return None
        content = json_str.strip()
        m = re.search(r"```(?:json)?\s*(.*?)\s*```", content, re.DOTALL)
        if m:
            content = m.group(1).strip()
        if content.lower().startswith("json"):
            content = re.sub(r"^json\s+", "", content, flags=re.IGNORECASE).strip()
        for parser in (json.loads, json5.loads, ast.literal_eval):
            try:
                return parser(content)
            except Exception:
                pass
        return None

    def requirement_classify(self, requirement: str) -> Dict[str, Any]:
        """Stage 2 — Requirement Classify (direct vs interpretive / category splits)."""
        system_prompt = self.prompts.get("requirement_classify", self.language)
        resp = self.runtime.generate(
            [{"role": "system", "content": system_prompt}, {"role": "user", "content": requirement}],
            stats_kind="analysis",
        )
        return self._json_parse(self._extract_content(resp)) or {}

    def directly_analysis(self, requirement_obj: Dict[str, Any], *, branch: str | None = None) -> Dict[str, Any]:
        """Stage 1 — Requirement Expansion for direct (explicit) requirements."""
        system_prompt = self.prompts.resolve_prompt("directly_analysis", self.language, branch)
        if not system_prompt:
            return {}
        resp = self.runtime.generate(
            [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"direct_requirement: {requirement_obj}"},
            ],
            stats_kind="analysis",
        )
        return self._json_parse(self._extract_content(resp)) or {}

    def interpretive_analysis(self, requirement_obj: Dict[str, Any], *, branch: str | None = None) -> Dict[str, Any]:
        """Stage 1 — Requirement Expansion for interpretive (content-aware) requirements."""
        system_prompt = self.prompts.resolve_prompt("interpretive_analysis", self.language, branch)
        if not system_prompt:
            return {}
        resp = self.runtime.generate(
            [{"role": "system", "content": system_prompt}, {"role": "user", "content": str(requirement_obj)}],
            stats_kind="analysis",
        )
        return self._json_parse(self._extract_content(resp)) or {}
