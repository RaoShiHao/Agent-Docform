# DocFormFlow Workflow Package

Orchestrates the four-stage formatting method:

1. **Requirement Expansion** — `AnalysisStage.directly_analysis` / `interpretive_analysis`
2. **Requirement Classify** — host routing + `AnalysisStage.requirement_classify`
3. **Target Element Localization** — `ClassificationStage`
4. **Verified Format Modification** — `DomainWorkflowPipeline.execute` (funcall → tools → check)

Entry points:

- CLI / API: `python -m workflow.run` or `from workflow.run import run_case`
- Host: `workflow/pipelines/host_pipeline.py`
- Domain: `workflow/pipelines/domain_pipeline.py`

See project [README.md](../README.md) / [README_zh.md](../README_zh.md).
