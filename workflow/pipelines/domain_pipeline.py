from __future__ import annotations

import copy
import hashlib
import json
import os
from typing import Any, Dict, List, Tuple

import win32com.client as win32

from tool.reader.image_reader import ImageReader
from tool.reader.page_reader import PageReader
from tool.reader.table_reader import TableReader
from tool.reader.text_reader import TextReader
from tool.word_com import create_word_app, open_document, release_word
from workflow.config.domain_config import load_domain_config
from workflow.core.types import StepResult, WorkflowContext
from workflow.llm.runtime import LLMRuntime
from workflow.prompts.prompt_hub import PromptHub
from workflow.stages.analysis import AnalysisStage
from workflow.stages.classification import ClassificationStage
from workflow.tools.action_executor import ActionExecutor
from workflow.tools.action_validator import ActionValidator
from workflow.tools.dynamic_registry import DynamicToolRegistry


class DomainWorkflowPipeline:
    """
    Per-domain DocFormFlow pipeline (page / text / table / image).

    Four stages (paper method):
      1) Requirement Expansion          — ``analyze`` → direct / interpretive analysis
      2) Requirement Classify           — ``analyze`` → ``requirement_classify`` (direct vs interpretive)
      3) Target Element Localization    — ``analyze`` → ClassificationStage (content → indices)
      4) Verified Format Modification   — ``execute`` → funcall → tools → read-back / check / refine

    Layers: config -> prompts -> llm -> stages -> executor.
    """

    def __init__(
        self,
        context: WorkflowContext,
        domain_name: str,
        config_path: str,
        funcall_prompt_key: str,
        check_prompt_key: str = "format_check",
    ) -> None:
        self.context = context
        self.domain_name = domain_name
        self.cfg = load_domain_config(config_path)
        self.prompts = PromptHub(self.cfg.prompt_config_path)
        self.runtime = LLMRuntime(self.cfg.llm_config, self.cfg.vllm_config)
        self.registry = DynamicToolRegistry(self.cfg.tools_config)
        # Current scope excludes character-level classification/editing for:
        # - text paragraph-character classification
        # - table character classification
        if self.domain_name in ("text", "table"):
            self.registry.tools.pop("CharacterTools", None)
        self.executor = ActionExecutor(self.registry)
        self.analysis_stage = AnalysisStage(self.runtime, self.prompts, context.language)
        self.classify_stage = ClassificationStage(
            self.runtime, self.prompts, context.language, self.cfg.understand_config
        )
        self.funcall_prompt_key = funcall_prompt_key
        self.check_prompt_key = check_prompt_key
        self.max_retry = int((self.cfg.refine_config or {}).get("max_retry", 0))
        self.max_check = int((self.cfg.refine_config or {}).get("max_check", 0))
        self.if_retry = bool((self.cfg.refine_config or {}).get("if_retry", False))
        self.if_check = bool((self.cfg.refine_config or {}).get("if_check", False))
        self.allowed_tools = self._resolve_allowed_tools()
        self.validator = ActionValidator(self.registry, self.allowed_tools)
        self.page_reader = PageReader()
        self.text_reader = TextReader()
        self.table_reader = TableReader()
        self.image_reader = ImageReader()
        self.function_params_map = self._build_function_param_map()
        self.cache_dir = os.path.join(self.context.save_dir, ".cache")
        os.makedirs(self.cache_dir, exist_ok=True)

    def _resolve_allowed_tools(self) -> set[str]:
        if self.domain_name == "page":
            return {"PageTools"}
        if self.domain_name == "text":
            return {"TextTools"}
        if self.domain_name == "table":
            return {"TableTools", "TextTools"}
        if self.domain_name == "image":
            return {"ImageTools"}
        return set(self.registry.tools.keys())

    def _build_function_param_map(self) -> Dict[str, Dict[str, str]]:
        return {
            "PageTools": {
                "set_margin": "margin",
                "set_gutter": "gutter",
                "set_paper": "paper",
                "set_grid": "grid",
                "set_columns": "columns",
                "set_footer_header_layout": "header_footer_layout",
                "set_footer_content": "footer_content",
                "set_header_content": "header_content",
            },
            "TextTools": {
                "set_base_font": "base_font",
                "set_advanced_font": "advanced_font",
                "set_outlinelevel": "outlinelevel",
                "set_alignment": "alignment",
                "set_pagination_control": "pagination_control",
                "set_spacing": "spacing",
                "set_indent": "indent",
                "set_partial_base_font": "partial_base_font",
                "set_partial_advanced_font": "partial_advanced_font",
            },
            "TableTools": {
                "set_table_width": "table_width",
                "set_table_text_wrapping": "text_wrapping",
                "set_table_pagination": "pagination",
                "set_table_alignment": "alignment",
                "set_table_left_indent": "left_indent",
            },
            "ImageTools": {
                "set_size": "size",
                "set_pagination": "pagination",
                "set_alignment": "alignment",
            },
        }

    @staticmethod
    def _remove_character_ops(obj: Any) -> Any:
        """
        Remove character-level requirements from nested structures.
        This keeps the pipeline focused on paragraph/table/image/page operations.
        """
        if isinstance(obj, dict):
            cleaned = {}
            for k, v in obj.items():
                lk = str(k).lower()
                if "character" in lk:
                    continue
                cleaned[k] = DomainWorkflowPipeline._remove_character_ops(v)
            return cleaned
        if isinstance(obj, list):
            return [DomainWorkflowPipeline._remove_character_ops(x) for x in obj]
        return obj

    @staticmethod
    def _extract_content(response: Dict[str, Any]) -> str:
        data = response.get("data") or {}
        return data.get("content", "") if response.get("success") else ""

    @staticmethod
    def _json_parse(raw: Any):
        import ast
        import json
        import re

        if isinstance(raw, (dict, list)):
            return raw
        if not isinstance(raw, str):
            return None
        content = raw.strip()
        m = re.search(r"```(?:json)?\s*(.*?)\s*```", content, re.DOTALL)
        if m:
            content = m.group(1).strip()
        if content.lower().startswith("json"):
            content = re.sub(r"^json\s+", "", content, flags=re.IGNORECASE).strip()
        parsers = [json.loads, ast.literal_eval]
        try:
            import json5  # type: ignore
            parsers.insert(1, json5.loads)
        except Exception:
            pass
        for parser in parsers:
            try:
                return parser(content)
            except Exception:
                pass
        return None

    @staticmethod
    def _stable_payload(data: Any) -> str:
        try:
            return json.dumps(data, ensure_ascii=False, sort_keys=True, default=str)
        except Exception:
            return str(data)

    def _cache_path(self, prefix: str, payload: Any) -> str:
        digest = hashlib.sha1(self._stable_payload(payload).encode("utf-8")).hexdigest()
        return os.path.join(self.cache_dir, f"{prefix}_{digest}.json")

    def _cache_get(self, prefix: str, payload: Any) -> Any:
        p = self._cache_path(prefix, payload)
        if not os.path.exists(p):
            return None
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _cache_set(self, prefix: str, payload: Any, value: Any) -> None:
        p = self._cache_path(prefix, payload)
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(value, f, ensure_ascii=False, indent=2, default=str)
        except Exception:
            pass

    @staticmethod
    def _prepare_classifier_user_content(domain_name: str, raw: str) -> str:
        r = (raw or "").strip()
        if not r:
            return ""
        if domain_name == "page":
            if r.startswith("{") and "page_requirement" in r:
                return raw
            return json.dumps({"page_requirement": [r]}, ensure_ascii=False)
        if domain_name in ("text", "table", "image"):
            if r.lstrip().startswith("["):
                return raw
            return json.dumps([{"instruction": r, "recognition_cues": "无"}], ensure_ascii=False)
        return raw

    @staticmethod
    def _split_requirement_classify(cls: Any, domain_name: str) -> tuple[Any, Any]:
        if not isinstance(cls, dict):
            return [], []
        if domain_name == "text":
            para = cls.get("paragraph")
            if isinstance(para, dict):
                return para.get("direct") or [], para.get("interpretive") or []
            return [], []
        if domain_name == "table":
            tbl = cls.get("table") if isinstance(cls.get("table"), dict) else {}
            cell = cls.get("cell") if isinstance(cls.get("cell"), dict) else {}
            direct_bundle = {
                "table_direct": tbl.get("direct") or [],
                "cell_direct": cell.get("direct") or [],
            }
            interp_bundle = {
                "table_interp": tbl.get("interpretive") or [],
                "cell_interp": cell.get("interpretive") or [],
            }
            return direct_bundle, interp_bundle
        return cls.get("direct") or [], cls.get("interpretive") or []

    def _run_post_classify_analysis(self, direct_raw: Any, interpretive_raw: Any) -> tuple[Dict[str, Any], Dict[str, Any]]:
        direct: Dict[str, Any] = {}
        interpretive: Dict[str, Any] = {}
        dn = self.domain_name

        if dn == "page":
            if isinstance(direct_raw, list) and direct_raw:
                out = self.analysis_stage.directly_analysis({"direct_requirement": direct_raw})
                if out:
                    direct = out
            if isinstance(interpretive_raw, list) and interpretive_raw:
                out = self.analysis_stage.interpretive_analysis({"page_requirement": interpretive_raw})
                if out:
                    interpretive = out
            return direct, interpretive

        if dn == "text":
            if isinstance(direct_raw, list) and direct_raw:
                out = self.analysis_stage.directly_analysis({"paragraph": direct_raw, "character": []})
                if out:
                    direct = out
            if isinstance(interpretive_raw, list) and interpretive_raw:
                out = self.analysis_stage.interpretive_analysis({"paragraph": interpretive_raw, "character": []})
                if out:
                    interpretive = out
            return direct, interpretive

        if dn == "image":
            if isinstance(direct_raw, list) and direct_raw:
                out = self.analysis_stage.directly_analysis({"direct_requirement": direct_raw})
                if out:
                    direct = out
            if isinstance(interpretive_raw, list) and interpretive_raw:
                out = self.analysis_stage.interpretive_analysis({"image": interpretive_raw})
                if out:
                    interpretive = out
            return direct, interpretive

        if dn == "table":
            if not isinstance(direct_raw, dict) or not isinstance(interpretive_raw, dict):
                return {}, {}
            td, cd = direct_raw.get("table_direct") or [], direct_raw.get("cell_direct") or []
            ti, ci = interpretive_raw.get("table_interp") or [], interpretive_raw.get("cell_interp") or []
            merged_direct: Dict[str, Any] = {}
            if td:
                out = self.analysis_stage.directly_analysis({"direct_requirement": td}, branch="table")
                if isinstance(out.get("tables"), list):
                    merged_direct["tables"] = out["tables"]
            if cd:
                out = self.analysis_stage.directly_analysis({"direct_requirement": cd}, branch="cell")
                if isinstance(out.get("cells"), list):
                    merged_direct["cells"] = out["cells"]
            merged_interp: Dict[str, Any] = {}
            if ti:
                out = self.analysis_stage.interpretive_analysis({"interpretive_requirement": ti}, branch="table")
                styles = out.get("table_styles")
                if isinstance(styles, list):
                    merged_interp["table"] = styles
            if ci:
                out = self.analysis_stage.interpretive_analysis({"interpretive_requirement": ci}, branch="cell")
                styles = out.get("cell_styles")
                if isinstance(styles, list):
                    merged_interp["cell"] = styles
            return merged_direct, merged_interp

        return {}, {}

    def _stringify_prompt_node(self, node: Any, mode: str | None = None) -> str:
        """
        Flatten nested yaml prompts such as ``page_funcall: { direct: { zh: str }, interpretive: { zh: str } }``
        into a single string for the current language.
        """
        lang = self.context.language
        cur: Any = node
        if mode and isinstance(cur, dict) and mode in cur:
            cur = cur[mode]
        while isinstance(cur, dict):
            if isinstance(cur.get(lang), str):
                return cur[lang]
            if isinstance(cur.get("zh"), str):
                return cur["zh"]
            if isinstance(cur.get("en"), str):
                return cur["en"]
            if len(cur) == 1:
                cur = next(iter(cur.values()))
            else:
                break
        return cur.strip() if isinstance(cur, str) else ""

    def _no_executable_work(self, direct: Any, interpretive: Any) -> bool:
        """True when this domain has nothing to run (skip Word / tools)."""

        def non_empty_list(x: Any) -> bool:
            return isinstance(x, list) and len(x) > 0

        if not direct and not interpretive:
            return True
        dn = self.domain_name
        if dn == "page":
            ds = isinstance(direct, dict) and non_empty_list(direct.get("sections"))
            ins = isinstance(interpretive, dict) and non_empty_list(interpretive.get("sections"))
            return not ds and not ins
        if dn == "text":
            if isinstance(direct, dict):
                for p in direct.get("paragraph") or []:
                    if isinstance(p, dict) and non_empty_list(p.get("requirement")):
                        return False
            if isinstance(interpretive, dict) and interpretive.get("style_structure"):
                return False
            return True
        if dn == "table":
            if isinstance(interpretive, dict):
                if non_empty_list(interpretive.get("table")) or non_empty_list(interpretive.get("cell")):
                    return False
            if isinstance(direct, dict):
                if non_empty_list(direct.get("tables")) or non_empty_list(direct.get("cells")):
                    return False
                tbl, cel = direct.get("table"), direct.get("cell")
                if isinstance(tbl, dict) and isinstance(cel, dict):
                    if any(
                        [
                            non_empty_list(tbl.get("direct")),
                            non_empty_list(tbl.get("interpretive")),
                            non_empty_list(cel.get("direct")),
                            non_empty_list(cel.get("interpretive")),
                        ]
                    ):
                        return False
            return True
        if dn == "image":
            if isinstance(direct, list) and len(direct) > 0:
                return False
            if isinstance(interpretive, dict):
                if non_empty_list(interpretive.get("image_structure")):
                    return False
            return True
        return False

    def _funcall_once(self, requirement_obj: Any, mode: str = "direct") -> List[Dict[str, Any]]:
        cache_payload = {
            "domain": self.domain_name,
            "mode": mode,
            "language": self.context.language,
            "requirement_obj": requirement_obj,
        }
        cached = self._cache_get("funcall", cache_payload)
        if isinstance(cached, list):
            return cached
        tool_desc = self.registry.describe(self.context.language)
        raw_prompt = self.prompts.get(self.funcall_prompt_key, self.context.language)
        system_prompt = self._stringify_prompt_node(raw_prompt, mode=mode)
        system_prompt = (system_prompt or "").replace("{tools_description}", tool_desc)
        try:
            resp = self.runtime.funcall_generate(
                [{"role": "system", "content": system_prompt}, {"role": "user", "content": str(requirement_obj)}]
            )
            parsed = self._json_parse(self._extract_content(resp)) or []
            if isinstance(parsed, list) and parsed:
                self._cache_set("funcall", cache_payload, parsed)
            return parsed
        except Exception:
            return cached if isinstance(cached, list) else []

    def _funcall_refine_once(
        self,
        requirement_obj: Any,
        previous_status: Any,
        last_funcall: Any,
    ) -> List[Dict[str, Any]]:
        cache_payload = {
            "domain": self.domain_name,
            "language": self.context.language,
            "requirement_obj": requirement_obj,
            "previous_status": previous_status,
            "last_funcall": last_funcall,
        }
        cached = self._cache_get("funcall_refine", cache_payload)
        if isinstance(cached, list):
            return cached
        tool_desc = self.registry.describe(self.context.language)
        raw_refine = self.prompts.get("funcall_refine", self.context.language)
        refine_prompt = self._stringify_prompt_node(raw_refine, mode=None)
        refine_prompt = (refine_prompt or "").replace("{tools_description}", tool_desc)
        user_content = (
            f"##requirement:\n{requirement_obj}\n"
            f"##last_funcall:\n{last_funcall}\n"
            f"##execution_status:\n{previous_status}\n"
            f"##refine_funcall:"
        )
        try:
            resp = self.runtime.funcall_generate(
                [{"role": "system", "content": refine_prompt}, {"role": "user", "content": user_content}]
            )
            parsed = self._json_parse(self._extract_content(resp)) or []
            if isinstance(parsed, list) and parsed:
                self._cache_set("funcall_refine", cache_payload, parsed)
            return parsed
        except Exception:
            return cached if isinstance(cached, list) else []

    def _format_check(self, requirement_obj: Any, state_obj: Any) -> Dict[str, Any]:
        cache_payload = {
            "domain": self.domain_name,
            "language": self.context.language,
            "requirement_obj": requirement_obj,
            "state_obj": state_obj,
        }
        cached = self._cache_get("format_check", cache_payload)
        if isinstance(cached, dict):
            return cached
        raw_check = self.prompts.get(self.check_prompt_key, self.context.language)
        prompt = self._stringify_prompt_node(raw_check, mode=None)
        try:
            resp = self.runtime.generate(
                [
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": f"##requirement:\n{requirement_obj}\n##result:\n{state_obj}"},
                ],
                stats_kind="check",
            )
            parsed = self._json_parse(self._extract_content(resp)) or {}
            if isinstance(parsed, dict) and parsed:
                self._cache_set("format_check", cache_payload, parsed)
            return parsed
        except Exception:
            return cached if isinstance(cached, dict) else {}

    @staticmethod
    def _extract_style_requirements(data: Any) -> List[Dict[str, Any]]:
        results: List[Dict[str, Any]] = []

        def _walk(obj: Any) -> None:
            if isinstance(obj, dict):
                if "style_name" in obj and "requirement" in obj:
                    results.append({"style_name": obj["style_name"], "requirement": obj["requirement"]})
                for v in obj.values():
                    _walk(v)
            elif isinstance(obj, list):
                for item in obj:
                    _walk(item)

        _walk(data)
        return results

    def _collect_reader_params(self, actions: List[Dict[str, Any]]) -> List[str]:
        params = []
        for action in actions:
            tool_name = action.get("tool_name")
            fn = action.get("function")
            key = (self.function_params_map.get(tool_name) or {}).get(fn)
            if key and key not in params:
                params.append(key)
        return params

    def _readback_state(self, doc, actions: List[Dict[str, Any]], args_replace: Dict[str, Any]) -> Any:
        params = self._collect_reader_params(actions)
        if self.domain_name == "page":
            section_list = args_replace.get("section_list", [1])
            sec = section_list[0] if isinstance(section_list, list) and section_list else 1
            if sec == "all":
                sec = 1
            return self.page_reader.read_page_properties(doc, int(sec), params_list=params, language=self.context.language)
        if self.domain_name == "text":
            loc_list = args_replace.get("location_list", [1])
            idx = loc_list[0] if isinstance(loc_list, list) and loc_list else 1
            if idx == "all":
                idx = 1
            return self.text_reader.read_text_properties(doc, int(idx), params_list=params, language=self.context.language)
        if self.domain_name == "table":
            if any(a.get("tool_name") == "TextTools" for a in actions):
                loc_list = args_replace.get("location_list", [1])
                idx = loc_list[0] if isinstance(loc_list, list) and loc_list else 1
                if idx == "all":
                    idx = 1
                return self.text_reader.read_text_properties(doc, int(idx), params_list=params, language=self.context.language)
            table_list = args_replace.get("table_list", [1])
            idx = table_list[0] if isinstance(table_list, list) and table_list else 1
            if idx == "all":
                idx = 1
            return self.table_reader.read_table_properties(doc, int(idx), params_list=params, language=self.context.language)
        if self.domain_name == "image":
            image_list = args_replace.get("image_list", [1])
            idx = image_list[0] if isinstance(image_list, list) and image_list else 1
            if idx == "all":
                idx = 1
            return self.image_reader.read_image_properties(doc, int(idx), params_list=params, language=self.context.language)
        return {}

    def analyze(self) -> StepResult:
        """
        Stages 1–3 for this domain:
          2) Requirement Classify  — direct vs interpretive (and domain-specific splits)
          1) Requirement Expansion — directly_analysis / interpretive_analysis
          3) Target Element Localization — ClassificationStage (when interpretive targets exist)
        """
        cache_payload = {
            "domain": self.domain_name,
            "language": self.context.language,
            "doc_path": self.context.doc_path,
            "requirement": self.context.requirement,
        }
        cached = self._cache_get("analyze", cache_payload)
        if isinstance(cached, dict):
            return StepResult(ok=True, step=f"{self.domain_name}.analyze", data=cached)
        if not (self.context.requirement or "").strip():
            empty: Dict[str, Any] = {"direct": {}, "interpretive": {}, "skipped": True}
            self._cache_set("analyze", cache_payload, empty)
            return StepResult(ok=True, step=f"{self.domain_name}.analyze", data=empty)
        try:
            # Stage 2 — Requirement Classify
            user_cls_in = self._prepare_classifier_user_content(self.domain_name, self.context.requirement)
            cls = self.analysis_stage.requirement_classify(user_cls_in)
            direct_raw, interpretive_raw = self._split_requirement_classify(cls, self.domain_name)
            # Stage 1 — Requirement Expansion (normalize into executable requirement objects)
            direct, interpretive = self._run_post_classify_analysis(direct_raw, interpretive_raw)
            if self.domain_name in ("text", "table"):
                direct = self._remove_character_ops(direct)
                interpretive = self._remove_character_ops(interpretive)
            payload: Dict[str, Any] = {"direct": direct, "interpretive": interpretive}
            # Stage 3 — Target Element Localization (map style/labels → document indices)
            if self.domain_name == "page" and interpretive.get("sections"):
                section_labels = interpretive.get("sections", [])
                payload["section_labels"] = self.classify_stage.classify_page_sections(
                    section_labels=copy.deepcopy(section_labels),
                    doc_path=self.context.doc_path,
                    save_dir=self.context.save_dir,
                )
            elif self.domain_name == "text" and interpretive.get("style_structure"):
                payload["text_labels"] = self.classify_stage.classify_text_paragraphs(
                    interpretive_styles=copy.deepcopy(interpretive),
                    doc_path=self.context.doc_path,
                    save_dir=self.context.save_dir,
                )
            elif self.domain_name == "table" and interpretive and (
                interpretive.get("table") or interpretive.get("cell")
            ):
                payload["table_labels"] = self.classify_stage.classify_table_elements(
                    interpretive_styles=copy.deepcopy(interpretive),
                    doc_path=self.context.doc_path,
                    save_dir=self.context.save_dir,
                )
            elif self.domain_name == "image" and interpretive.get("image_structure"):
                payload["image_labels"] = self.classify_stage.classify_images(
                    interpretive_styles=copy.deepcopy(interpretive),
                    doc_path=self.context.doc_path,
                    save_dir=self.context.save_dir,
                )
            self._cache_set("analyze", cache_payload, payload)
            return StepResult(ok=True, step=f"{self.domain_name}.analyze", data=payload)
        except Exception as e:
            if isinstance(cached, dict):
                return StepResult(ok=True, step=f"{self.domain_name}.analyze", data=cached)
            return StepResult(ok=False, step=f"{self.domain_name}.analyze", error=str(e))

    @staticmethod
    def _status_has_execution_error(status: Any) -> bool:
        if not isinstance(status, list):
            return False
        for x in status:
            if not isinstance(x, dict):
                continue
            if x.get("status") == "error":
                return True
            for v in x.values():
                if isinstance(v, dict) and v.get("status") == "error":
                    return True
        return False

    def _execute_group(self, doc, requirement_obj: Any, args_replace: Dict[str, Any], mode: str = "direct") -> Dict[str, Any]:
        """
        Stage 4 helper — Verified Format Modification for one requirement group:
          funcall → validate → execute tools → (optional) read-back / format_check / refine.
        """
        last_status: Any = None
        last_actions: Any = None
        current_req = copy.deepcopy(requirement_obj)
        for check_round in range(0, self.max_check + 1):
            raw_actions = self._funcall_once(current_req, mode=mode)
            if not raw_actions:
                return {
                    "status": "error",
                    "code": "EMPTY_FUNCALL",
                    "message": "empty funcall",
                    "check_round": check_round,
                    "raw_actions": raw_actions,
                }

            last_raw = raw_actions
            val_feedback: Any = None
            actions: List[Dict[str, Any]] = []
            validation_errors: List[Any] = []
            for vk in range(0, self.max_retry + 1):
                if vk > 0:
                    if not self.if_retry:
                        break
                    raw_actions = self._funcall_refine_once(current_req, val_feedback, last_raw)
                    if not raw_actions:
                        return {
                            "status": "error",
                            "code": "EMPTY_FUNCALL_REFINE",
                            "message": "refine funcall returned empty after validation failure",
                            "check_round": check_round,
                            "validation_round": vk,
                            "last_validation_feedback": val_feedback,
                        }
                    last_raw = raw_actions

                actions, validation_errors = self.validator.validate_actions(raw_actions)
                if not validation_errors:
                    break

                val_feedback = {
                    "phase": "action_validation",
                    "message": (
                        "The previous function-call JSON was rejected before execution "
                        "(wrong tool name, disallowed tool, missing function, bad parameters, etc.). "
                        "Fix tool_name to match registry (e.g. PageTools not PageTool), parameters, and schema."
                    ),
                    "errors": [e.__dict__ for e in validation_errors],
                }
                last_raw = raw_actions
                if not self.if_retry or vk >= self.max_retry:
                    return {
                        "status": "error",
                        "code": "ACTION_VALIDATION_FAILED",
                        "message": "action validation failed after refine retries" if vk > 0 else "action validation failed",
                        "check_round": check_round,
                        "validation_rounds": vk,
                        "raw_actions": raw_actions,
                        "validation_errors": [e.__dict__ for e in validation_errors],
                    }

            status = self.executor.execute(actions, args_replace=args_replace)
            last_status = status
            last_actions = actions
            if self.if_retry:
                exec_refines = 0
                while exec_refines < self.max_retry and self._status_has_execution_error(status):
                    retried_actions = self._funcall_refine_once(current_req, status, last_actions)
                    if not retried_actions:
                        break
                    actions2, errors2 = self.validator.validate_actions(retried_actions)
                    if not actions2 or errors2:
                        break
                    last_actions = actions2
                    status = self.executor.execute(actions2, args_replace=args_replace)
                    last_status = status
                    exec_refines += 1

            if not self.if_check:
                return {
                    "status": "success",
                    "actions": last_actions,
                    "execute_status": last_status,
                }
            state_obj = self._readback_state(doc, last_actions or actions, args_replace=args_replace)
            check_result = self._format_check(current_req, state_obj)
            if check_result.get("satisfy"):
                return {
                    "status": "success",
                    "actions": last_actions,
                    "execute_status": last_status,
                    "readback": state_obj,
                    "check": check_result,
                }
            if not check_result.get("new_command"):
                return {
                    "status": "error",
                    "actions": last_actions,
                    "execute_status": last_status,
                    "readback": state_obj,
                    "check": check_result,
                }
            current_req = check_result.get("new_command")
        return {"status": "error", "execute_status": last_status, "message": "max_check_reached"}

    def _build_interpretive_tasks(self, analysis_payload: Dict[str, Any], interpretive_req: Any) -> List[Tuple[Any, Dict[str, Any]]]:
        tasks: List[Tuple[Any, Dict[str, Any]]] = []
        if self.domain_name == "page":
            labels = analysis_payload.get("section_labels") or []
            name_to_idx: Dict[str, Any] = {}
            for item in labels:
                if isinstance(item, dict) and item.get("section_name") is not None:
                    name_to_idx[str(item["section_name"])] = item.get("section_index")

            if isinstance(interpretive_req, dict):
                sections = interpretive_req.get("sections")
                if isinstance(sections, list) and sections:
                    for sec in sections:
                        if not isinstance(sec, dict):
                            continue
                        sname = sec.get("section_name")
                        if sname is None:
                            continue
                        idx = name_to_idx.get(str(sname))
                        if idx is None:
                            continue
                        single_req = {"sections": [copy.deepcopy(sec)]}
                        tasks.append((single_req, {"doc": None, "section_list": [idx]}))
            if not tasks and isinstance(interpretive_req, dict) and interpretive_req.get("sections"):
                args_replace: Dict[str, Any] = {"doc": None}
                secs = interpretive_req.get("sections")
                if isinstance(secs, list) and secs and labels:
                    sn = secs[0].get("section_name") if isinstance(secs[0], dict) else None
                    if sn is not None:
                        for item in labels:
                            if isinstance(item, dict) and item.get("section_name") == sn:
                                args_replace["section_list"] = [item.get("section_index")]
                                break
                tasks.append((interpretive_req, args_replace))
            return tasks

        if self.domain_name == "text":
            paragraph_labels = (analysis_payload.get("text_labels") or {}).get("paragraph_labels", {})
            for req in self._extract_style_requirements(interpretive_req):
                style = req.get("style_name")
                location_list = paragraph_labels.get(style, [])
                if location_list:
                    tasks.append((req.get("requirement"), {"doc": None, "location_list": location_list}))
            return tasks

        if self.domain_name == "table":
            labels = analysis_payload.get("table_labels") or {}
            table_labels = labels.get("table_labels", {})
            cell_labels = labels.get("cell_labels", {})
            if isinstance(interpretive_req, dict):
                for mode, req_list in interpretive_req.items():
                    if not isinstance(req_list, list):
                        continue
                    for req in req_list:
                        style = req.get("style_name")
                        if not style:
                            continue
                        if mode == "table":
                            table_list = table_labels.get(style, [])
                            if table_list:
                                tasks.append((req.get("requirement"), {"doc": None, "table_list": table_list}))
                        elif mode == "cell":
                            location_list = cell_labels.get(style, [])
                            if location_list:
                                tasks.append((req.get("requirement"), {"doc": None, "location_list": location_list}))
            return tasks

        if self.domain_name == "image":
            image_labels = (analysis_payload.get("image_labels") or {}).get("image_labels", {})
            for req in self._extract_style_requirements(interpretive_req):
                style = req.get("style_name")
                image_list = image_labels.get(style, [])
                if image_list:
                    tasks.append((req.get("requirement"), {"doc": None, "image_list": image_list}))
            return tasks

        return tasks

    def execute(self, analysis_payload: Dict[str, Any]) -> StepResult:
        """
        Stage 4 — Verified Format Modification.

        Apply direct and interpretive formatting via Word COM tools, with optional
        validation, execution refine, and format-check loops.
        """
        direct_req = analysis_payload.get("direct")
        interpretive_req = analysis_payload.get("interpretive")
        cache_payload = {
            "domain": self.domain_name,
            "language": self.context.language,
            "doc_path": self.context.doc_path,
            "analysis_payload": analysis_payload,
        }
        try:
            if analysis_payload.get("skipped"):
                return StepResult(
                    ok=True,
                    step=f"{self.domain_name}.execute",
                    data={"skipped": True, "reason": "no_domain_requirement"},
                )
            if self._no_executable_work(direct_req, interpretive_req):
                return StepResult(
                    ok=True,
                    step=f"{self.domain_name}.execute",
                    data={"skipped": True, "reason": "no_formatting_requirement_for_domain"},
                )
            word = create_word_app(visible=False)
            doc = open_document(word, self.context.doc_path)
            results = {"direct": None, "interpretive": None}
            if direct_req:
                results["direct"] = self._execute_group(doc, direct_req, {"doc": doc}, mode="direct")
            if interpretive_req:
                tasks = self._build_interpretive_tasks(analysis_payload, interpretive_req)
                task_results = []
                if not tasks:
                    task_results.append({"status": "skipped", "reason": "no_interpretive_targets"})
                for req_obj, args_replace in tasks:
                    args_replace["doc"] = doc
                    task_results.append(
                        self._execute_group(doc, req_obj, args_replace, mode="interpretive")
                    )
                results["interpretive"] = task_results
            self._cache_set("execute", cache_payload, results)
            return StepResult(ok=True, step=f"{self.domain_name}.execute", data=results)
        except Exception as e:
            cached = self._cache_get("execute", cache_payload)
            if isinstance(cached, dict):
                return StepResult(ok=True, step=f"{self.domain_name}.execute", data=cached)
            return StepResult(ok=False, step=f"{self.domain_name}.execute", error=str(e))
        finally:
            release_word(
                word if "word" in locals() else None,
                doc if "doc" in locals() else None,
                save_changes=True,
            )

