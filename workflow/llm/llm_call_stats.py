from __future__ import annotations

import threading
from typing import Any, Dict, Mapping, Optional

_lock = threading.Lock()
_by_kind: Dict[str, int] = {
    "analysis": 0,
    "check": 0,
    "classify": 0,
    "funcall": 0,
    "unspecified": 0,
}
_total: int = 0
_input_tokens: int = 0
_output_tokens: int = 0
_total_tokens: int = 0


def reset_llm_call_stats() -> None:
    """Clear counters (call once at the start of a workflow run)."""
    global _total, _input_tokens, _output_tokens, _total_tokens
    with _lock:
        for k in list(_by_kind.keys()):
            _by_kind[k] = 0
        _total = 0
        _input_tokens = 0
        _output_tokens = 0
        _total_tokens = 0


def record_llm_call(kind: str) -> int:
    """Increment counters for one LLM API invocation routed through ``LLMRuntime``.

    Returns the 1-based call index for this run (useful for progress logs).
    """
    global _total
    with _lock:
        _by_kind[kind] = _by_kind.get(kind, 0) + 1
        _total += 1
        return _total


def record_llm_usage(usage: Optional[Mapping[str, Any]]) -> None:
    """Accumulate token usage returned by one successful LLM response."""
    global _input_tokens, _output_tokens, _total_tokens
    if not usage:
        return
    prompt = int(usage.get("prompt_tokens") or usage.get("input_tokens") or 0)
    completion = int(usage.get("completion_tokens") or usage.get("output_tokens") or 0)
    total = int(usage.get("total_tokens") or (prompt + completion))
    with _lock:
        _input_tokens += prompt
        _output_tokens += completion
        _total_tokens += total


def get_llm_call_stats() -> Dict[str, Any]:
    """Snapshot of counts (thread-safe copy)."""
    with _lock:
        return {
            "total_calls": _total,
            "by_kind": dict(_by_kind),
        }


def get_token_use() -> Dict[str, Any]:
    """Snapshot for ``token_use.json`` (thread-safe copy)."""
    with _lock:
        calls = _total
        avg = round(_total_tokens / calls, 2) if calls else 0.0
        return {
            "input_tokens": _input_tokens,
            "output_tokens": _output_tokens,
            "total_tokens": _total_tokens,
            "llm_calls": calls,
            "avg_tokens_per_call": avg,
        }
