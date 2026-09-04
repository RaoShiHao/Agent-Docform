from __future__ import annotations

import argparse
import os

from workflow.pipelines.host_pipeline import HostWorkflowPipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Minimal workflow test sample")
    parser.add_argument("--doc", required=True, help="Source.docx path")
    parser.add_argument("--requirement", required=True, help="Formatting requirement text")
    parser.add_argument("--save-dir", required=True, help="Output directory")
    parser.add_argument("--language", default="zh", choices=["zh", "en"])
    args = parser.parse_args()

    os.makedirs(args.save_dir, exist_ok=True)
    pipeline = HostWorkflowPipeline(save_dir=args.save_dir, language=args.language)
    report = pipeline.run(requirement=args.requirement, source_doc=args.doc)
    print("Workflow finished.")
    print(f"Report: {os.path.join(args.save_dir, 'workflow_report.json')}")
    print(f"Modified doc: {os.path.join(args.save_dir, 'modified.docx')}")
    print(f"Domain keys: {list((report.get('domains') or {}).keys())}")


if __name__ == "__main__":
    main()
