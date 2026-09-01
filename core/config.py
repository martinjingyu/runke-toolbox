"""加载 config.yaml，本机专属的覆盖项从 config.local.yaml 合并进来。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config() -> dict[str, Any]:
    with open(PROJECT_ROOT / "config.yaml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    local_path = PROJECT_ROOT / "config.local.yaml"
    if local_path.exists():
        with open(local_path, "r", encoding="utf-8") as f:
            local_config = yaml.safe_load(f) or {}
        config = _merge(config, local_config)

    return config
