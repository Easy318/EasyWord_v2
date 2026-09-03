"""进程入口：供开发 `python -m word` 与 PyInstaller 冻结共用。"""

from __future__ import annotations

import multiprocessing
import os
import sys
from pathlib import Path


def _ensure_stdio() -> None:
    """windowed 冻结（无控制台）时 stdout/stderr 可能为 None，uvicorn 写日志会崩。"""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def main() -> None:
    _ensure_stdio()
    multiprocessing.freeze_support()

    import uvicorn

    from word.config.settings import get_host, get_port

    host = get_host()
    port = get_port()

    if _is_frozen():
        # 打包产物：传入 app 对象，避免字符串模块路径解析问题
        from word.app import app

        uvicorn.run(app, host=host, port=port, log_level="info")
        return

    # 开发态：默认热重载
    reload_dir = str(Path(__file__).resolve().parent)
    uvicorn.run(
        "word.app:app",
        host=host,
        port=port,
        reload=True,
        reload_dirs=[reload_dir],
        log_level="info",
    )


if __name__ == "__main__":
    main()
