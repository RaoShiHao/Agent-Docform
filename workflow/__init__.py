from workflow.pipelines.host_pipeline import HostWorkflowPipeline


def run_case(*args, **kwargs):
    from workflow.run import run_case as _run_case

    return _run_case(*args, **kwargs)


__all__ = [
    "HostWorkflowPipeline",
    "run_case",
]
