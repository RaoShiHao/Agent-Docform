from __future__ import annotations

import os
import shutil
from typing import Any, Dict, List, Optional, Tuple

from tool.file_trans import FileConverter
from workflow.config.domain_config import load_domain_config
from workflow.core.types import WorkflowContext
from workflow.llm.runtime import LLMRuntime
from workflow.prompts.prompt_hub import PromptHub
from workflow.pipelines.domain_pipeline import DomainWorkflowPipeline
from workflow.llm.llm_call_stats import get_llm_call_stats, get_token_use, reset_llm_call_stats
from workflow.stages.analysis import AnalysisStage


# DocFormFlow four-stage workflow (paper / method overview)
# ------------------------------------------------------------
# 1) Requirement Expansion  — expand / normalize the user requirement
# 2) Requirement Classify   — split requirements by domain (page/text/table/image)
#                              and by direct vs interpretive style
# 3) Target Element Localization — locate concrete targets in the document
# 4) Verified Format Modification — generate tool calls, execute, read back, verify
# ------------------------------------------------------------
# Host owns stage (2) at the document level (domain routing).
# Each DomainWorkflowPipeline then runs stages (1)–(4) within its domain.


class HostWorkflowPipeline:
    """
    Host pipeline: route a full-document requirement into domain sub-requirements,
    then run domain pipelines in order: page -> text -> table -> image.

    Stage mapping at host level:
      - Requirement Classify: ``_route_requirements`` / ``AnalysisStage.requirement_classify``

    ``agent_config_dir`` selects which YAML pack to use (default ``config/app_agent``,
    or e.g. ``experiment/configs/gemini-3-flash`` for batch experiments).
    """

    def __init__(
        self,
        save_dir: str,
        language: str = "zh",
        *,
        agent_config_dir: Optional[str] = None,
        host_config_path: Optional[str] = None,
    ) -> None:
        self.save_dir = save_dir
        self.language = language if language in ("zh", "en") else "zh"
        self.agent_config_dir = (agent_config_dir or "config/app_agent").replace("\\", "/").rstrip("/")
        self.host_config_path = host_config_path or f"{self.agent_config_dir}/host_agent.yaml"
        self.file_tool = FileConverter()
        os.makedirs(save_dir, exist_ok=True)

    def _route_requirements(self, requirement: str) -> Dict[str, str]:
        """
        Stage 2 — Requirement Classify (host / domain routing).

        Split the user requirement by category (Page/Text/Table/Image) using host_prompt.
        """
        empty = {"page": "", "text": "", "table": "", "image": ""}
        req = (requirement or "").strip()
        if not req:
            return empty
        parts: Dict[str, List[str]] = {k: [] for k in empty}
        cat_map = {
            "page": "page",
            "text": "text",
            "table": "table",
            "image": "image",
            "页面": "page",
            "文本": "text",
            "表格": "table",
            "图片": "image",
            "页面类": "page",
            "文本类": "text",
            "表格类": "table",
            "图片类": "image",
        }
        try:
            cfg = load_domain_config(self.host_config_path)
            runtime = LLMRuntime(cfg.llm_config, cfg.vllm_config)
            hub = PromptHub(cfg.prompt_config_path)
            stage = AnalysisStage(runtime, hub, self.language)
            raw = stage.requirement_classify(req)
            instructions = raw.get("instructions") if isinstance(raw, dict) else None
            if not isinstance(instructions, list):
                # Do NOT broadcast the full requirement to every domain — that
                # pollutes Image/Table with Text/Page instructions.
                print(
                    "[Host] requirement_classify returned no instructions list; "
                    "leaving all domain routes empty.",
                    flush=True,
                )
                return empty.copy()
            for item in instructions:
                if not isinstance(item, dict):
                    continue
                cat_raw = str(item.get("category", "")).strip()
                cat_key = cat_raw.split("|")[0].strip()
                dom = cat_map.get(cat_key) or cat_map.get(cat_key.lower()) or cat_map.get(cat_raw) or cat_map.get(cat_raw.lower())
                if not dom:
                    continue
                instr = item.get("instruction")
                if isinstance(instr, str) and instr.strip():
                    parts[dom].append(instr.strip())
            routes = {k: "\n".join(parts[k]) for k in empty}
            if not any(routes.values()):
                print(
                    "[Host] requirement_classify produced no routed instructions; "
                    "leaving all domain routes empty.",
                    flush=True,
                )
                return empty.copy()
            # Guard: Host sometimes pastes the entire requirement into several
            # categories. That reintroduces cross-domain pollution.
            full_hits = [k for k, v in routes.items() if (v or "").strip() == req]
            if len(full_hits) >= 2:
                print(
                    f"[Host] domains {full_hits} each received the full unsplit "
                    "requirement; clearing those routes to avoid cross-domain pollution.",
                    flush=True,
                )
                for k in full_hits:
                    routes[k] = ""
                if not any(routes.values()):
                    return empty.copy()
            return routes
        except Exception as e:
            print(
                f"[Host] requirement routing failed ({e}); "
                "leaving all domain routes empty (no full-req broadcast).",
                flush=True,
            )
            return empty.copy()

    def _domain_specs(self) -> List[Tuple[str, str, str]]:
        d = self.agent_config_dir
        return [
            ("page", f"{d}/page_agent.yaml", "page_funcall"),
            ("text", f"{d}/text_agent.yaml", "text_funcall"),
            ("table", f"{d}/table_agent.yaml", "table_funcall"),
            ("image", f"{d}/image_agent.yaml", "image_funcall"),
        ]

    def _working_doc(self, source_doc: str) -> str:
        dst = os.path.join(self.save_dir, "modified.docx")
        shutil.copy(source_doc, dst)
        return dst

    def run(self, requirement: str, source_doc: str) -> Dict[str, Any]:
        """
        End-to-end DocFormFlow run:
          Host Requirement Classify → per-domain stages (1)–(4).
        """
        reset_llm_call_stats()
        doc_path = self._working_doc(source_doc)
        # Stage 2 (host): Requirement Classify → domain routes
        routes = self._route_requirements(requirement)
        report: Dict[str, Any] = {
            "language": self.language,
            "source_doc": source_doc,
            "working_doc": doc_path,
            "agent_config_dir": self.agent_config_dir,
            "host_routes": routes,
            "domains": {},
        }
        try:
            for domain, cfg, funcall_key in self._domain_specs():
                domain_save_dir = os.path.join(self.save_dir, domain.capitalize())
                os.makedirs(domain_save_dir, exist_ok=True)
                sub_req = routes.get(domain, "")
                ctx = WorkflowContext(
                    requirement=sub_req,
                    doc_path=doc_path,
                    save_dir=domain_save_dir,
                    language=self.language,
                )
                pipeline = DomainWorkflowPipeline(
                    context=ctx,
                    domain_name=domain,
                    config_path=cfg,
                    funcall_prompt_key=funcall_key,
                    check_prompt_key="format_check",
                )
                # Domain stages 1–3 live in analyze(); stage 4 in execute()
                analyze = pipeline.analyze()
                execute = pipeline.execute(analyze.data if analyze.ok else {})
                report["domains"][domain] = {"analyze": analyze.__dict__, "execute": execute.__dict__}

            report["llm_call_stats"] = get_llm_call_stats()
            self.file_tool.write_json_file(report, os.path.join(self.save_dir, "workflow_report.json"))
            return report
        finally:
            self.file_tool.write_json_file(get_token_use(), os.path.join(self.save_dir, "token_use.json"))
