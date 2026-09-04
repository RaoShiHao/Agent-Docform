from __future__ import annotations

import re
import sys
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

# Running ``workflow/run.py`` often puts ``.../workflow`` first on ``sys.path``, so bare
# ``import llm`` wrongly loads this package (``workflow.llm``) instead of project-root
# ``llm`` → circular import. Put repo root first so ``llm.openai_client`` resolves correctly.
_root_s = str(Path(__file__).resolve().parents[2])
try:
    sys.path.remove(_root_s)
except ValueError:
    pass
sys.path.insert(0, _root_s)

from llm.openai_client import OpenaiLLMClient, OpenaiLLMImageClient
from workflow.llm.llm_call_stats import record_llm_call, record_llm_usage

# Gateway / transport flakes: sleep then retry (covers 500 DB, 503 channel, timeouts, etc.).
_LLM_RETRY_SLEEP_SEC = 10.0
_LLM_RETRY_MAX_ATTEMPTS = 5

_RETRYABLE_PATTERNS = (
    r"\b500\b",
    r"\b502\b",
    r"\b503\b",
    r"\b504\b",
    r"connection refused",
    r"connection reset",
    r"connect(?:ion)? timed? ?out",
    r"request timed out",
    r"timed? ?out",
    r"temporarily unavailable",
    r"model_not_found",
    r"no available channel",
    r"query_data_error",
    r"new_api_error",
    r"database error",
    r"dial tcp",
    r"server error",
    r"overloaded",
    r"rate limit",
    r"too many requests",
    r"\b429\b",
)


def _is_ready_llm_config(cfg: Dict[str, Any] | None) -> bool:
    if not cfg:
        return False
    for key in ("model", "api_key", "base_url"):
        val = str(cfg.get(key) or "").strip()
        if not val or val.startswith("${"):
            return False
    return True


def _record_response_usage(response: Dict[str, Any]) -> None:
    if not response or not response.get("success"):
        return
    data = response.get("data") or {}
    record_llm_usage(data.get("usage"))


def _error_text(response: Dict[str, Any] | None = None, exc: BaseException | None = None) -> str:
    parts: List[str] = []
    if exc is not None:
        parts.append(str(exc))
    if response:
        err = response.get("error") or {}
        if isinstance(err, dict):
            parts.append(str(err.get("message") or ""))
            parts.append(str(err.get("code") or ""))
            parts.append(str(err.get("type") or ""))
        else:
            parts.append(str(err))
    return " ".join(parts).lower()


def _is_retryable_llm_failure(
    response: Dict[str, Any] | None = None,
    exc: BaseException | None = None,
) -> bool:
    text = _error_text(response, exc)
    if not text.strip():
        return False
    for pat in _RETRYABLE_PATTERNS:
        if re.search(pat, text, flags=re.IGNORECASE):
            return True
    return False


def _log_llm_begin(call_no: int, kind: str, *, vision: bool = False, n_images: int = 0) -> None:
    extra = f", vision images={n_images}" if vision else ""
    print(f"[LLM] #{call_no} {kind} requesting{extra} ...", flush=True)


def _log_llm_end(call_no: int, kind: str, response: Dict[str, Any], elapsed: float) -> None:
    ok = bool(response and response.get("success"))
    if ok:
        usage = (response.get("data") or {}).get("usage") or {}
        tokens = usage.get("total_tokens", 0)
        print(f"[LLM] #{call_no} {kind} done in {elapsed:.1f}s (ok, tokens={tokens})", flush=True)
    else:
        err = ((response or {}).get("error") or {}).get("message") or "unknown error"
        print(f"[LLM] #{call_no} {kind} done in {elapsed:.1f}s (fail: {err})", flush=True)


def _invoke_with_retry(
    *,
    call_no: int,
    kind: str,
    invoke: Callable[[], Dict[str, Any]],
) -> Dict[str, Any]:
    """Call LLM; on retryable gateway/transport errors sleep and retry."""
    t0 = time.perf_counter()
    last_response: Dict[str, Any] = {
        "success": False,
        "error": {"code": 500, "message": "LLM call failed before any attempt"},
    }
    for attempt in range(1, _LLM_RETRY_MAX_ATTEMPTS + 1):
        try:
            response = invoke()
        except Exception as e:
            elapsed = time.perf_counter() - t0
            if attempt < _LLM_RETRY_MAX_ATTEMPTS and _is_retryable_llm_failure(exc=e):
                print(
                    f"[LLM] #{call_no} {kind} attempt {attempt}/{_LLM_RETRY_MAX_ATTEMPTS} "
                    f"exception (retryable): {e}; sleep {_LLM_RETRY_SLEEP_SEC:.0f}s then retry",
                    flush=True,
                )
                time.sleep(_LLM_RETRY_SLEEP_SEC)
                continue
            print(
                f"[LLM] #{call_no} {kind} done in {elapsed:.1f}s (exception: {e})",
                flush=True,
            )
            raise

        last_response = response if isinstance(response, dict) else last_response
        if response and response.get("success"):
            _record_response_usage(response)
            _log_llm_end(call_no, kind, response, time.perf_counter() - t0)
            return response

        err = ((response or {}).get("error") or {}).get("message") or "unknown error"
        if attempt < _LLM_RETRY_MAX_ATTEMPTS and _is_retryable_llm_failure(response=response):
            print(
                f"[LLM] #{call_no} {kind} attempt {attempt}/{_LLM_RETRY_MAX_ATTEMPTS} "
                f"fail (retryable): {err}; sleep {_LLM_RETRY_SLEEP_SEC:.0f}s then retry",
                flush=True,
            )
            time.sleep(_LLM_RETRY_SLEEP_SEC)
            continue

        _log_llm_end(call_no, kind, response, time.perf_counter() - t0)
        return response

    _log_llm_end(call_no, kind, last_response, time.perf_counter() - t0)
    return last_response


class LLMRuntime:
    def __init__(self, llm_config: Dict[str, Any], vllm_config: Dict[str, Any] | None = None) -> None:
        self.llm = OpenaiLLMClient(llm_config)
        self._vllm_config = vllm_config or {}
        self._vllm: Optional[OpenaiLLMImageClient] = None
        if _is_ready_llm_config(self._vllm_config):
            self._vllm = OpenaiLLMImageClient(self._vllm_config)

    @property
    def vllm(self) -> Optional[OpenaiLLMImageClient]:
        return self._vllm

    def generate(
        self,
        messages: List[Dict[str, str]],
        use_vision: bool = False,
        image_inputs: List[str] | None = None,
        *,
        stats_kind: str = "unspecified",
    ) -> Dict[str, Any]:
        call_no = record_llm_call(stats_kind)
        images = image_inputs or []
        _log_llm_begin(call_no, stats_kind, vision=bool(use_vision), n_images=len(images))

        def _invoke() -> Dict[str, Any]:
            if use_vision:
                if self._vllm is None:
                    if not _is_ready_llm_config(self._vllm_config):
                        raise ValueError(
                            "Vision is requested but vllm_config is incomplete. "
                            "Set DOCFORMFLOW_VLLM_* env vars or fill vllm_config in yaml."
                        )
                    self._vllm = OpenaiLLMImageClient(self._vllm_config)
                return self._vllm.generate(messages, image_inputs=images)
            return self.llm.generate(messages)

        return _invoke_with_retry(call_no=call_no, kind=stats_kind, invoke=_invoke)

    def funcall_generate(self, messages: List[Dict[str, str]]) -> Dict[str, Any]:
        call_no = record_llm_call("funcall")
        _log_llm_begin(call_no, "funcall")
        return _invoke_with_retry(
            call_no=call_no,
            kind="funcall",
            invoke=lambda: self.llm.funcall_generate(messages),
        )
