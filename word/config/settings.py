"""EasyWord 运行时配置（本机 sidecar，仅监听 loopback）。"""

from __future__ import annotations

import os

from word.config.constants import DEFAULT_HOST, DEFAULT_PORT


def get_host() -> str:
    return os.environ.get("EASYWORD_HOST", DEFAULT_HOST)


def get_port() -> int:
    raw = os.environ.get("EASYWORD_PORT", str(DEFAULT_PORT))
    try:
        return int(raw)
    except ValueError:
        return DEFAULT_PORT
