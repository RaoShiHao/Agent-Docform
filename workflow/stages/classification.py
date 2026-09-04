from __future__ import annotations

import os
from typing import Any, Dict, List

from tool.understanding.doc_understand import DocInform
from workflow.llm.runtime import LLMRuntime
from workflow.prompts.prompt_hub import PromptHub


class ClassificationStage:
    """
    Stage 3 — Target Element Localization.

    Map interpretive style / section / image labels onto concrete document indices
    (paragraphs, tables, cells, images, sections), optionally with vision context.
    """

    def __init__(self, runtime: LLMRuntime, prompts: PromptHub, language: str, understand_config: Dict[str, Any]) -> None:
        self.runtime = runtime
        self.prompts = prompts
        self.language = language
        self.understand_config = understand_config or {}
        self.doc_inform = DocInform()

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
        for parser in (json.loads, json5.loads, ast.literal_eval):
            try:
                return parser(content)
            except Exception:
                pass
        return None

    def classify_page_sections(self, section_labels: List[Dict[str, Any]], doc_path: str, save_dir: str) -> List[Dict[str, Any]]:
        page_infos = self.doc_inform.info_get(
            doc_path,
            save_dir=os.path.join(save_dir, "understand"),
            scope="page",
            content_x=self.understand_config.get("page_limit"),
            is_vision=self.understand_config.get("is_vision"),
        )
        prompt = self.prompts.get("section_classify", self.language)
        result: List[Dict[str, Any]] = []
        for page_info in page_infos:
            messages = [
                {"role": "system", "content": prompt},
                {
                    "role": "user",
                    "content": f"section_name_dict: {section_labels}\nsection_info={page_info}",
                },
            ]
            use_vision = bool(self.understand_config.get("is_vision"))
            image_inputs = []
            if use_vision:
                start = page_info["section_range"].get("start")
                end = page_info["section_range"].get("end")
                end = min(start + (self.understand_config.get("page_limit") or 1), end)
                image_inputs = [os.path.join(save_dir, f"understand/doc_images/page_{x}.png") for x in range(start, end)]
            resp = self.runtime.generate(
                messages, use_vision=use_vision, image_inputs=image_inputs, stats_kind="classify"
            )
            parsed = self._json_parse(self._extract_content(resp))
            if parsed:
                result.append(parsed)
        return result

    @staticmethod
    def _remove_key_inplace(data: Any, remove_key: str) -> Any:
        if isinstance(data, dict):
            if remove_key in data:
                del data[remove_key]
            for _, value in data.items():
                ClassificationStage._remove_key_inplace(value, remove_key)
        elif isinstance(data, list):
            for item in data:
                ClassificationStage._remove_key_inplace(item, remove_key)
        return data

    @staticmethod
    def _labels_sum(items: List[Dict[str, Any]], name_key: str, index_key: str) -> Dict[str, List[int]]:
        out: Dict[str, List[int]] = {}
        for item in items:
            style_name = item.get(name_key)
            idx = item.get(index_key)
            if not style_name or idx is None:
                continue
            out.setdefault(style_name, []).append(idx)
        return out

    def classify_text_paragraphs(self, interpretive_styles: Dict[str, Any], doc_path: str, save_dir: str) -> Dict[str, Any]:
        paragraph_labels = (interpretive_styles or {}).get("style_structure")
        if not paragraph_labels:
            return {"paragraph_labels": {}}
        page_text_infos = self.doc_inform.info_get(
            doc_path,
            save_dir=os.path.join(save_dir, "understand"),
            scope="text",
            is_vision=self.understand_config.get("is_vision"),
        )
        style_dict = self._remove_key_inplace(paragraph_labels, "requirement")
        prompt = self.prompts.get("paragraph_classify", self.language)
        classify_result: List[Dict[str, Any]] = []
        for page_index, page_text_info in page_text_infos.items():
            messages = [
                {"role": "system", "content": prompt},
                {
                    "role": "user",
                    "content": f"## input:\n### style_dict:\n{style_dict}\n###paragraphs:\n{page_text_info}\n##output:",
                },
            ]
            use_vision = bool(self.understand_config.get("is_vision"))
            image_inputs = []
            if use_vision:
                image_inputs = [os.path.join(save_dir, f"understand/doc_images/page_{page_index}.png")]
            resp = self.runtime.generate(
                messages, use_vision=use_vision, image_inputs=image_inputs, stats_kind="classify"
            )
            parsed = self._json_parse(self._extract_content(resp))
            if isinstance(parsed, list):
                classify_result.extend(parsed)
        return {"paragraph_labels": self._labels_sum(classify_result, "style_name", "index")}

    @staticmethod
    def _table_to_matrix(cells: List[Dict[str, Any]]) -> List[List[str]]:
        if not cells:
            return []
        max_row = max((c.get("row", 0) for c in cells), default=0)
        max_col = max((c.get("col", 0) for c in cells), default=0)
        table = [["" for _ in range(max_col + 1)] for _ in range(max_row + 1)]
        for cell in cells:
            r = cell.get("row")
            c = cell.get("col")
            if isinstance(r, int) and isinstance(c, int) and r >= 0 and c >= 0:
                table[r][c] = cell.get("paragraph", "")
        return table

    def classify_table_elements(self, interpretive_styles: Dict[str, Any], doc_path: str, save_dir: str) -> Dict[str, Any]:
        table_styles = (interpretive_styles or {}).get("table")
        cell_styles = (interpretive_styles or {}).get("cell")
        if not table_styles and not cell_styles:
            return {"table_labels": {}, "cell_labels": {}}
        table_infos = self.doc_inform.info_get(
            doc_path,
            save_dir=os.path.join(save_dir, "understand"),
            scope="table",
            before=self.understand_config.get("before", 0),
            after=self.understand_config.get("after", 0),
            is_vision=self.understand_config.get("is_vision"),
        )
        table_prompt = self.prompts.get("table_classify", self.language)
        cell_prompt = self.prompts.get("cell_classify", self.language)
        table_out: List[Dict[str, Any]] = []
        cell_out: List[Dict[str, Any]] = []
        table_style_list = self._remove_key_inplace(table_styles, "requirement") if table_styles else []
        cell_style_list = self._remove_key_inplace(cell_styles, "requirement") if cell_styles else []

        for table_info in table_infos:
            table_index = table_info.get("table_index")
            page_range = table_info.get("page_range", {})
            start_page = page_range.get("start_page", 1)
            end_page = page_range.get("end_page", start_page)
            use_vision = bool(self.understand_config.get("is_vision"))
            image_inputs = []
            if use_vision:
                image_inputs = [
                    os.path.join(save_dir, f"understand/doc_images/page_{p}.png")
                    for p in range(start_page, end_page + 1)
                ]

            if table_style_list:
                table_payload = {
                    "table_index": table_index,
                    "content": self._table_to_matrix(table_info.get("cells", [])),
                    "before_texts": table_info.get("before_texts"),
                    "after_texts": table_info.get("after_texts"),
                }
                messages = [
                    {"role": "system", "content": table_prompt},
                    {
                        "role": "user",
                        "content": f"## input:\n### style_list:\n{table_style_list}\n###table_info:\n{table_payload}\n##output:",
                    },
                ]
                resp = self.runtime.generate(
                    messages, use_vision=use_vision, image_inputs=image_inputs, stats_kind="classify"
                )
                parsed = self._json_parse(self._extract_content(resp))
                if isinstance(parsed, dict):
                    table_out.append(parsed)

            if cell_style_list:
                messages = [
                    {"role": "system", "content": cell_prompt},
                    {
                        "role": "user",
                        "content": f"## input:\n### style_list:\n{cell_style_list}\n###cells_info:\n{table_info.get('cells', [])}\n##output:",
                    },
                ]
                resp = self.runtime.generate(
                    messages, use_vision=use_vision, image_inputs=image_inputs, stats_kind="classify"
                )
                parsed = self._json_parse(self._extract_content(resp))
                if isinstance(parsed, list):
                    cell_out.extend(parsed)
        return {
            "table_labels": self._labels_sum(table_out, "style_name", "table_index"),
            "cell_labels": self._labels_sum(cell_out, "style_name", "index"),
        }

    def classify_images(self, interpretive_styles: Dict[str, Any], doc_path: str, save_dir: str) -> Dict[str, Any]:
        image_structure = (interpretive_styles or {}).get("image_structure")
        if not image_structure:
            return {"image_labels": {}}
        style_dict = self._remove_key_inplace(image_structure, "requirement")
        image_infos = self.doc_inform.info_get(
            doc_path,
            save_dir=os.path.join(save_dir, "understand"),
            scope="image",
            before=self.understand_config.get("before", 0),
            after=self.understand_config.get("after", 0),
            is_vision=self.understand_config.get("is_vision"),
        )
        prompt = self.prompts.get("image_classify", self.language)
        image_out: List[Dict[str, Any]] = []
        for image_info in image_infos:
            page_number = image_info.get("page_number")
            use_vision = bool(self.understand_config.get("is_vision"))
            image_inputs = []
            if use_vision and page_number:
                image_inputs = [os.path.join(save_dir, f"understand/doc_images/page_{page_number}.png")]
            messages = [
                {"role": "system", "content": prompt},
                {
                    "role": "user",
                    "content": f"image_name_dict={style_dict}\nimage_info={image_info}",
                },
            ]
            resp = self.runtime.generate(
                messages, use_vision=use_vision, image_inputs=image_inputs, stats_kind="classify"
            )
            parsed = self._json_parse(self._extract_content(resp))
            if isinstance(parsed, dict):
                image_out.append(parsed)
        return {"image_labels": self._labels_sum(image_out, "image_name", "image_index")}

