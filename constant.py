from pathlib import Path

ABS_DIR = Path(__file__).resolve().parent


def load_dotenv(path: str | None = None) -> None:
    """Load simple KEY=VALUE pairs from a `.env` file into ``os.environ`` (no override)."""
    import os

    env_path = Path(path) if path else ABS_DIR / ".env"
    if not env_path.is_file():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value
