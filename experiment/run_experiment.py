"""
Batch experiment runner for DocFormFlow (Full method).

Edit the ``# ===== run settings =====`` block below, then::

    python experiment/run_experiment.py

Internal subprocess mode (do not call manually)::

    python experiment/run_experiment.py --run-one <index>
"""

from __future__ import annotations

import os
import sys
import time
import subprocess
from pathlib import Path

# Ensure project root is on sys.path when launched as ``python experiment/run_experiment.py``
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from constant import ABS_DIR, load_dotenv
from tool.file_trans import FileConverter
from workflow.run import run_case

# ===== run settings (edit here) =====
CONFIG_DIR = "experiment/configs/gemini-3-flash"
DATA_ROOT = r"path\to\DocFormBench"          # dataset root: <DATA_ROOT>/<index>/init.docx
SAVE_ROOT = r"path\to\DocFormFlow_outputs"   # output root
SCOPE_PREFIX = ""  # optional path prefix under SAVE_ROOT
START = 1
END = 601
STEP = 1  # use negative step to count down, e.g. START=500, END=250, STEP=-1
REQUIREMENT_FILE = "requirement_hard.md"  # fixed; no easy/medium level switch
# ===== end settings =====


class ExperimentRun:
    def __init__(self) -> None:
        self.config_dir = CONFIG_DIR.replace("\\", "/").rstrip("/")
        self.data_root = DATA_ROOT
        self.save_root = SAVE_ROOT
        self.filetool = FileConverter()
        models_info = Path(self.config_dir).name
        prefix = f"{SCOPE_PREFIX.strip('/')}/" if SCOPE_PREFIX.strip() else ""
        # No R-hard (or other level) directory layer.
        self.scope_name = f"{prefix}{models_info}"

    def read_md_table(self, md_path: str) -> str:
        with open(md_path, "r", encoding="utf-8") as f:
            return f.read()

    def _sample_paths(self, index: int) -> tuple[str, str, str]:
        data_sample_dir = os.path.join(self.data_root, str(index))
        save_dir = os.path.join(self.save_root, self.scope_name, str(index))
        status_file = os.path.join(save_dir, "execution_status.json")
        return data_sample_dir, save_dir, status_file

    def run_single_index(self, index: int) -> None:
        load_dotenv()
        data_sample_dir, save_dir, status_file = self._sample_paths(index)
        os.makedirs(save_dir, exist_ok=True)
        start_time = time.time()
        try:
            requirement = self.read_md_table(
                os.path.join(data_sample_dir, REQUIREMENT_FILE)
            )
            file_path = os.path.join(data_sample_dir, "init.docx")
            language_config = os.path.join(data_sample_dir, "meta_info.json")
            language = self.filetool.read_json_file(language_config).get("language") or "zh"
            if language not in ("zh", "en"):
                language = "zh"

            run_case(
                requirement=requirement,
                doc=file_path,
                save_dir=save_dir,
                language=language,
                agent_config_dir=self.config_dir,
            )
            run_time = time.time() - start_time
            self.filetool.write_json_file(
                data={
                    "status": "execution pass",
                    "time": run_time,
                    "config_dir": self.config_dir,
                    "scope_name": self.scope_name,
                    "index": index,
                },
                file_path=status_file,
            )
        except Exception as e:
            run_time = time.time() - start_time
            self.filetool.write_json_file(
                data={
                    "status": "execution error!",
                    "time": run_time,
                    "error": str(e),
                    "config_dir": self.config_dir,
                    "scope_name": self.scope_name,
                    "index": index,
                },
                file_path=status_file,
            )
            raise

    def datasets_run(self, start: int = START, end: int = END, step: int = STEP) -> None:
        if step == 0:
            raise ValueError("STEP must be non-zero")
        for index in range(start, end, step):
            _, save_dir, status_file = self._sample_paths(index)
            fail_file = os.path.join(save_dir, "TaskFail.json")
            if os.path.exists(status_file) or os.path.exists(fail_file):
                continue

            cmd = [
                sys.executable,
                str(Path(__file__).resolve()),
                "--run-one",
                str(index),
            ]
            try:
                subprocess.run(cmd, check=True, cwd=str(ABS_DIR))
            except subprocess.CalledProcessError:
                self.filetool.save_error(save_dir=save_dir)
                continue


if __name__ == "__main__":
    # Subprocess worker: only needs the sample index; all other knobs are in-script.
    if len(sys.argv) >= 3 and sys.argv[1] == "--run-one":
        ExperimentRun().run_single_index(int(sys.argv[2]))
        sys.exit(0)

    ExperimentRun().datasets_run(start=START, end=END, step=STEP)
