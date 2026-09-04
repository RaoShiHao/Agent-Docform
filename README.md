<<<<<<< HEAD
# DocFormFlow

**DocFormFlow** is a Windows-oriented document formatting workflow: given a natural-language formatting requirement and a `.docx` file, it uses an LLM plus Word COM tools to produce a revised document.

> Chinese README: [README_zh.md](README_zh.md)

## Method Overview (Four Stages)

DocFormFlow decouples **what to format** from **how to format**:

| Stage | Name | Role |
|------:|------|------|
| 1 | **Requirement Expansion** | Expand / normalize requirements into executable structures (`directly_analysis` / `interpretive_analysis`) |
| 2 | **Requirement Classify** | Split by domain (`page` / `text` / `table` / `image`) and by direct vs interpretive |
| 3 | **Target Element Localization** | Locate concrete targets in the document (paragraph / table / image / section indices) |
| 4 | **Verified Format Modification** | Generate tool calls → execute via Word COM → read back → check / refine |

Host pipeline runs stage 2 for domain routing, then each domain pipeline runs stages 1–4.

```
requirement + .docx
        │
        ▼
 HostWorkflowPipeline
  - copy → modified.docx
  - Stage 2: route → page / text / table / image
        │
        ▼
 DomainWorkflowPipeline (per domain)
  Stage 1 Expansion → Stage 2 Classify → Stage 3 Localization
        → Stage 4 Verified Modification (funcall / tools / check)
        │
        ▼
 modified.docx + workflow_report.json
```

## Requirements

| Item | Notes |
|------|--------|
| OS | **Windows** (Word COM via `pywin32`) |
| Microsoft Word | **Word 2016 or later** (desktop; licensed) — **required** |
| Python | 3.10+ recommended |
| LLM API | Any OpenAI-compatible chat API |

Optional: vision model when `understand_config.is_vision: true`; `zai` for GLM; Poppler for DOCX→PDF→images.

### Initialize Word COM constants (required)

```bash
python -m win32com.client.makepy
```

Select **Microsoft Word xx.0 Object Library**, then confirm.

## Quick Start

```bash
cd DocFormFlow
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m win32com.client.makepy

copy .env.example .env
# Edit .env: DOCFORMFLOW_MODEL / DOCFORMFLOW_API_KEY / DOCFORMFLOW_BASE_URL
# Optional vision: DOCFORMFLOW_VLLM_MODEL / DOCFORMFLOW_VLLM_API_KEY / DOCFORMFLOW_VLLM_BASE_URL
```

### Run a single formatting case

```bash
python -m workflow.run ^
  --requirement "Set body paragraphs to Songti Xiaosi with 1.5 line spacing" ^
  --doc "examples\docs\Word_test.docx" ^
  --save-dir "workspace\demo_run" ^
  --language en
```

Python API:

```python
from workflow.run import run_case

report = run_case(
    requirement="Set body paragraphs to Songti Xiaosi with 1.5 line spacing",
    doc=r"examples\docs\Word_test.docx",
    save_dir=r"workspace\demo_run",
    language="en",
    # agent_config_dir="experiment/configs/gemini-3-flash",  # optional YAML pack
)
```

Outputs under `--save-dir`:

- `modified.docx` — working copy after all domain edits
- `workflow_report.json` — routing / analyze / execute summary
- `token_use.json` — token usage
- `Page/`, `Text/`, `Table/`, `Image/` — per-domain artifacts and `.cache/`

## Configuration

Configs: `config/app_agent/*.yaml` (default) or `experiment/configs/<model>/`.

Secrets must use placeholders (never commit real keys):

| Variable | Purpose |
|----------|---------|
| `DOCFORMFLOW_MODEL` | Chat model name |
| `DOCFORMFLOW_API_KEY` | API key |
| `DOCFORMFLOW_BASE_URL` | OpenAI-compatible base URL |
| `DOCFORMFLOW_VLLM_*` | Optional vision model settings |

- Domain prompts: `config/prompt_config/`
- Tool schemas: `config/Tools/`

## Batch Experiments

See also [experiment/README.md](experiment/README.md).

1. Prepare DocFormBench-style data:

```
<DATA_ROOT>/
  <index>/
    init.docx
    requirement_hard.md
    meta_info.json          # e.g. {"language": "zh"}
```

2. Edit the settings block in `experiment/run_experiment.py`:

```python
CONFIG_DIR = "experiment/configs/gemini-3-flash"
DATA_ROOT = r"path\to\DocFormBench"
SAVE_ROOT = r"path\to\DocFormFlow_outputs"
START = 1
END = 601
STEP = 1
REQUIREMENT_FILE = "requirement_hard.md"
```

3. Ensure `.env` is filled, then run:

```bash
cd DocFormFlow
python experiment/run_experiment.py
```

Each sample is executed in a subprocess. Outputs go to:

`<SAVE_ROOT>/<model>/<index>/` (e.g. `modified.docx`, `workflow_report.json`, `execution_status.json`).

Already-finished samples (with `execution_status.json` or `TaskFail.json`) are skipped.

## Project Layout

```
DocFormFlow/
├── workflow/           # Host + domain pipelines (4 stages)
├── tool/               # Word COM tools & readers
├── llm/                # OpenAI-compatible clients
├── config/             # Default YAML configs
├── experiment/         # Batch runner + model config packs
├── examples/docs/      # Sample .docx
├── constant.py         # Project root + .env loader
├── requirements.txt
├── README.md
└── README_zh.md
```

## Notes

- Keep API keys out of git (use `.env` / environment variables).
- End-to-end formatting requires real Windows + Microsoft Word.

## License

Add a license file appropriate for your release before publishing.
=======
# Agent-Docform
>>>>>>> 40bb6ba479ea11660bdc28ed20d46c0b89feb4ed
