# Experiment configs & batch runner

## Config packs

Under `experiment/configs/<model>/`:

| File | Role |
|------|------|
| `host_agent.yaml` | Host routing LLM (Requirement Classify at document level) |
| `page_agent.yaml` / `text_agent.yaml` / `table_agent.yaml` / `image_agent.yaml` | Domain pipelines (stages 1–4) |

Bundled example:

- `gemini-3-flash` — Gemini Flash with vision enabled (`is_vision: True`)

Secrets use env placeholders (`DOCFORMFLOW_API_KEY`, `DOCFORMFLOW_BASE_URL`, `DOCFORMFLOW_VLLM_*`).  
Copy `.env.example` → `.env` at the project root and fill in values **before** running. Do not hard-code keys in YAML.

## Batch run

1. Edit the settings block at the top of `experiment/run_experiment.py`:

| Knob | Meaning |
|------|---------|
| `CONFIG_DIR` | Model YAML pack, e.g. `experiment/configs/gemini-3-flash` |
| `DATA_ROOT` | Dataset root (`<index>/init.docx`, `requirement_hard.md`, `meta_info.json`) |
| `SAVE_ROOT` | Output root |
| `START` / `END` / `STEP` | Sample index range (negative `STEP` supported) |
| `REQUIREMENT_FILE` | Fixed to `requirement_hard.md` |

2. Run:

```bash
cd DocFormFlow
python experiment/run_experiment.py
```

Save path layout:

`<SAVE_ROOT>/<model>/<index>/`

Already finished samples (`execution_status.json` or `TaskFail.json`) are skipped.

## Single-case run (non-batch)

```bash
python -m workflow.run ^
  --requirement "..." ^
  --doc "path\to\init.docx" ^
  --save-dir "workspace\demo" ^
  --language zh ^
  --agent-config-dir experiment/configs/gemini-3-flash
```
