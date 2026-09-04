from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Optional

from constant import load_dotenv
from workflow.pipelines.host_pipeline import HostWorkflowPipeline


def run_case(
    requirement: str,
    doc: str,
    save_dir: str,
    language: str = "zh",
    *,
    agent_config_dir: Optional[str] = None,
) -> dict[str, Any]:
    """Run one workflow case; same behavior as ``python -m workflow.run`` CLI."""
    load_dotenv()
    pipeline = HostWorkflowPipeline(
        save_dir=save_dir,
        language=language,
        agent_config_dir=agent_config_dir,
    )
    return pipeline.run(requirement=requirement, source_doc=doc)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="DocFormFlow: format a .docx document from a natural-language requirement."
    )
    parser.add_argument("--requirement", required=True, help="Formatting requirement text")
    parser.add_argument("--doc", required=True, help="Path to the source .docx file")
    parser.add_argument("--save-dir", required=True, help="Output directory for modified.docx and reports")
    parser.add_argument("--language", default="zh", choices=["zh", "en"])
    parser.add_argument(
        "--agent-config-dir",
        default=None,
        help="YAML pack dir relative to project root, e.g. experiment/configs/qwen3-vl-30b-a3b",
    )
    args = parser.parse_args()

    run_case(
        requirement=args.requirement,
        doc=args.doc,
        save_dir=args.save_dir,
        language=args.language,
        agent_config_dir=args.agent_config_dir,
    )


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        main()
    else:
        root = Path(__file__).resolve().parents[1]
        requirement = "将正文段落设置为宋体小四，1.5倍行距"
        doc = str(root / "examples" / "docs" / "Word_test.docx")
        save_dir = str(root / "workspace" / "demo_run")
        print(f"[DocFormFlow demo] doc={doc}")
        print(f"[DocFormFlow demo] save_dir={save_dir}")
        run_case(requirement, doc, save_dir, language="zh")
