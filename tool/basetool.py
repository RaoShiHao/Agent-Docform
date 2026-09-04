from functools import wraps
import yaml, os
from constant import ABS_DIR


class ToolMeta(type):
    """Metaclass: create an independent method registry for each subclass."""

    def __new__(cls, name, bases, attrs):
        # Initialize registry when creating the new class
        new_class = super().__new__(cls, name, bases, attrs)
        new_class._registry = []
        # Collect methods decorated as tools in the current class body
        for attr_name, attr_value in attrs.items():
            if callable(attr_value) and hasattr(attr_value, "__tool_metadata__"):
                new_class._registry.append(attr_value)
        return new_class

    def validate_metadata(cls, metadata):
        """Validate the structure of ``@register_tool`` metadata."""
        # Top level must contain both en and zh
        required_langs = {"en", "zh"}
        if missing := required_langs - metadata.keys():
            raise ValueError(f"Missing required language keys: {missing}")

        for lang in required_langs:
            lang_data = metadata[lang]
            if not isinstance(lang_data, dict):
                raise TypeError(f"Value for {lang} must be a dict")

            required_fields = {"function_description", "params"}
            if missing := required_fields - lang_data.keys():
                raise ValueError(f"Missing required fields in {lang}: {missing}")

            params = lang_data["params"]
            if not isinstance(params, dict):
                raise TypeError(f"params for {lang} must be a dict")

            for param_name, param_info in params.items():
                if not isinstance(param_info, dict):
                    raise TypeError(f"params['{param_name}'] in {lang} must be a dict")

                required_param_keys = {"type", "description"}
                if missing := required_param_keys - param_info.keys():
                    raise ValueError(f"params['{param_name}'] in {lang} missing fields: {missing}")


class BaseTool(metaclass=ToolMeta):
    """Base tool class; each subclass gets its own registry."""

    @classmethod
    def register_tool(cls, metadata):
        cls.validate_metadata(metadata=metadata)
        """Decorator factory: mark a method and attach bilingual metadata."""

        def decorator(func):
            func.__tool_metadata__ = metadata

            @wraps(func)
            def wrapper(self, *args, **kwargs):
                return func(self, *args, **kwargs)

            return wrapper

        return decorator

    @classmethod
    def get_tools(cls):
        """Return all registered tools for the current subclass."""
        return [
            {
                "name": func.__name__,
                "metadata": getattr(func, "__tool_metadata__", None),
            }
            for func in cls._registry
        ]

    @classmethod
    def describe_tools(cls, language="en"):
        """Build multilingual tool descriptions for LLM prompts."""
        if language not in ["zh", "en"]:
            language = "en"
        descriptions = []
        for tool in cls.get_tools():
            meta = tool["metadata"].get(language, {})
            desc = {
                "name": tool["name"],
                "description": meta.get("function_description", ""),
                "parameters": meta.get("params", {}),
            }
            descriptions.append(desc)
        return descriptions


class ContextToolsConfig:
    def __init__(self, config_path="/config/Tools/TextToolsConfig.yaml"):
        # Resolve to an absolute path under the project root
        abs_path = os.path.join(ABS_DIR, config_path.lstrip("/"))
        with open(abs_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)


# Example calculator tool (independent registry)
class Calculator(BaseTool):
    @BaseTool.register_tool(
        {
            "en": {
                "function_description": "Addition operation",
                "params": {
                    "a": {"type": "number", "description": "First operand"},
                    "b": {"type": "number", "description": "Second operand"},
                },
            },
            "zh": {
                "function_description": "加法运算",
                "params": {
                    "a": {"type": "数字", "description": "第一个操作数"},
                    "b": {"type": "数字", "description": "第二个操作数"},
                },
            },
        }
    )
    def add(self, a, b):
        return a + b


# Example text processor (independent registry)
class TextProcessor(BaseTool):
    @BaseTool.register_tool(
        {
            "en": {
                "function_description": "Text tokenization",
                "params": {"text": {"type": "string", "description": "Input text"}},
            },
            "zh": {
                "function_description": "文本分词",
                "params": {"text": {"type": "字符串", "description": "输入文本"}},
            },
        }
    )
    def tokenize(self, text):
        return text.split()


if __name__ == "__main__":
    print(Calculator.get_tools())
    # Output: [{'name': 'add', 'metadata': ...}]

    print(TextProcessor.get_tools())
    # Output: [{'name': 'tokenize', 'metadata': ...}]

    print(Calculator.describe_tools("zh"))
    # Output: [{'name': 'add', 'description': '加法运算', 'parameters': {...}}]
