from openai import OpenAI
import base64
import mimetypes
from typing import Dict, Any, List
import os
from constant import ABS_DIR


def _build_client(model: str, base_url: str, api_key: str):
    """Create an OpenAI-compatible client; use Zhipu SDK when model name contains ``glm``."""
    if "glm" in (model or "").lower():
        try:
            from zai import ZhipuAiClient
        except ImportError as e:
            raise ImportError(
                "Model name contains 'glm' but package 'zai' is not installed. "
                "Install with: pip install zai"
            ) from e
        return ZhipuAiClient(api_key=api_key)
    return OpenAI(base_url=base_url, api_key=api_key)


def _usage_from_response(response) -> Dict[str, int]:
    usage = getattr(response, "usage", None)
    if usage is None:
        return {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
    prompt = int(getattr(usage, "prompt_tokens", 0) or 0)
    completion = int(getattr(usage, "completion_tokens", 0) or 0)
    total = int(getattr(usage, "total_tokens", 0) or (prompt + completion))
    return {
        "prompt_tokens": prompt,
        "completion_tokens": completion,
        "total_tokens": total,
    }


class OpenaiLLMClient:
    def __init__(self, llm_config: Dict[str, Any]):
        """Initialize the client.
        :param llm_config: Must include the following parameters:
            model: Model identifier
            base_url: API endpoint
            api_key: API key
        """
        required_keys = {"model", "base_url", "api_key"}
        if missing := required_keys - llm_config.keys():
            raise ValueError(f"Missing required config parameters: {missing}")

        self.model = llm_config.get("model") or os.getenv("DOCFORMFLOW_MODEL", "")
        self.base_url = llm_config.get("base_url") or os.getenv("DOCFORMFLOW_BASE_URL", "")
        self.api_key = llm_config.get("api_key") or os.getenv("DOCFORMFLOW_API_KEY", "")
        if not self.model or not self.base_url or not self.api_key:
            raise ValueError(
                "LLM config incomplete. Set model/api_key/base_url in yaml "
                "or via DOCFORMFLOW_MODEL / DOCFORMFLOW_API_KEY / DOCFORMFLOW_BASE_URL."
            )
        if self.api_key.startswith("${") or self.model.startswith("${"):
            raise ValueError(
                "LLM placeholders were not expanded. Export DOCFORMFLOW_* env vars "
                "or replace ${...} values in config/app_agent/*.yaml."
            )

        self.llm_params = llm_config.get("params", {})
        self.params_config = {
            "temperature": self.llm_params.get("temperature", 0.6),
            "top_p": self.llm_params.get("top_p", 1.0),
            "seed": self.llm_params.get("seed", 42),
            "max_tokens": self.llm_params.get("max_tokens", 8128),
            **self.llm_params,
        }
        self.funcall_params_config = {
            "temperature": self.llm_params.get("temperature", 0.0),
            "top_p": self.llm_params.get("top_p", 1.0),
            "seed": self.llm_params.get("seed", 42),
            "max_tokens": self.llm_params.get("max_tokens", 8128),
            **self.llm_params,
        }
        self._client = _build_client(self.model, self.base_url, self.api_key)

    def close(self):
        if hasattr(self, "_client"):
            self._client.close()

    def generate(self, messages) -> Dict[str, Any]:
        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=messages,
                stream=False,
                **self.params_config,
            )
            return {
                "success": True,
                "data": {
                    "content": response.choices[0].message.content,
                    "model": response.model,
                    "usage": _usage_from_response(response),
                },
            }
        except Exception as e:
            return {
                "success": False,
                "error": {"code": 500, "message": str(e)},
            }

    def funcall_generate(self, messages) -> Dict[str, Any]:
        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=messages,
                stream=False,
                **self.funcall_params_config,
            )
            return {
                "success": True,
                "data": {
                    "content": response.choices[0].message.content,
                    "model": response.model,
                    "usage": _usage_from_response(response),
                },
            }
        except Exception as e:
            return {
                "success": False,
                "error": {"code": 500, "message": str(e)},
            }

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


class OpenaiLLMImageClient:
    def __init__(self, llm_config: Dict[str, Any]):
        required_keys = {"model", "base_url", "api_key"}
        if missing := required_keys - llm_config.keys():
            raise ValueError(f"Missing required config parameters: {missing}")

        self.model = llm_config.get("model") or os.getenv("DOCFORMFLOW_VLLM_MODEL", "")
        self.base_url = llm_config.get("base_url") or os.getenv("DOCFORMFLOW_VLLM_BASE_URL", "")
        self.api_key = llm_config.get("api_key") or os.getenv("DOCFORMFLOW_VLLM_API_KEY", "")
        if not self.model or not self.base_url or not self.api_key:
            raise ValueError(
                "Vision LLM config incomplete. Set vllm_config in yaml "
                "or via DOCFORMFLOW_VLLM_* environment variables."
            )
        if self.api_key.startswith("${") or self.model.startswith("${"):
            raise ValueError(
                "Vision LLM placeholders were not expanded. Export DOCFORMFLOW_VLLM_* "
                "or replace ${...} values in config/app_agent/*.yaml."
            )

        self.system_prompt = llm_config.get("system_prompt", "You are an image understanding assistant.")
        self.llm_params = llm_config.get("params", {})
        self.params_config = {
            "temperature": self.llm_params.get("temperature", 0.6),
            "top_p": self.llm_params.get("top_p", 1.0),
            "seed": self.llm_params.get("seed", 42),
            "max_tokens": self.llm_params.get("max_tokens", 8128),
            **self.llm_params,
        }
        self.funcall_params_config = {
            "temperature": self.llm_params.get("temperature", 0.0),
            "top_p": self.llm_params.get("top_p", 1.0),
            "seed": self.llm_params.get("seed", 42),
            "max_tokens": self.llm_params.get("max_tokens", 8128),
            **self.llm_params,
        }
        self._client = _build_client(self.model, self.base_url, self.api_key)

    def close(self):
        if hasattr(self, "_client"):
            self._client.close()

    def _is_url(self, path: str) -> bool:
        return path.startswith("http://") or path.startswith("https://")

    def _encode_image_to_base64(self, image_path: str) -> str:
        mime_type, _ = mimetypes.guess_type(image_path)
        if not mime_type or not mime_type.startswith("image/"):
            raise ValueError(f"Unable to determine image MIME type: {image_path}")

        with open(image_path, "rb") as f:
            image_data = f.read()
        base64_data = base64.b64encode(image_data).decode("utf-8")
        return f"data:{mime_type};base64,{base64_data}"

    def generate(self, messages: List[Dict[str, Any]], image_inputs: List[str] = None) -> Dict[str, Any]:
        try:
            processed_messages = messages.copy()

            if image_inputs and processed_messages:
                last_user_message_index = None
                for i in range(len(processed_messages) - 1, -1, -1):
                    if processed_messages[i]["role"] == "user":
                        last_user_message_index = i
                        break

                if last_user_message_index is not None:
                    last_user_msg = processed_messages[last_user_message_index]
                    if isinstance(last_user_msg["content"], str):
                        content: List[Dict[str, Any]] = [
                            {"type": "text", "text": last_user_msg["content"]}
                        ]
                        for image_path in image_inputs:
                            if self._is_url(image_path):
                                image_url = image_path
                            else:
                                image_url = self._encode_image_to_base64(image_path)
                            content.append(
                                {"type": "image_url", "image_url": {"url": image_url}}
                            )
                        processed_messages[last_user_message_index] = {
                            "role": "user",
                            "content": content,
                        }

            response = self._client.chat.completions.create(
                model=self.model,
                messages=processed_messages,
                stream=False,
                **self.params_config,
            )
            return {
                "success": True,
                "data": {
                    "content": response.choices[0].message.content,
                    "model": response.model,
                    "usage": _usage_from_response(response),
                },
            }
        except Exception as e:
            return {
                "success": False,
                "error": {"code": 500, "message": str(e)},
            }

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
