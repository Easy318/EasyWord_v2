"""路径规范化（用于文档绑定比较）。"""

from __future__ import annotations

import os
from pathlib import Path


def normalize_path(path: str) -> str:
    """绝对路径 + 规范化大小写/分隔符，便于 Documents 比对。"""
    p = Path(path).expanduser()
    try:
        p = p.resolve(strict=False)
    except OSError:
        p = Path(os.path.abspath(str(p)))
    return os.path.normcase(str(p))


def paths_equal(a: str, b: str) -> bool:
    """比较两路径是否同一文件（规范化 + samefile）。"""
    if normalize_path(a) == normalize_path(b):
        return True
    try:
        return os.path.samefile(a, b)
    except OSError:
        return False
