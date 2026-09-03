"""集中注册各 core 模块路由（对齐 QHub register_router）。"""

from __future__ import annotations

from fastapi import FastAPI

from word.core.content_control import content_control_api
from word.core.document import document_api
from word.core.health import health_api
from word.core.preview import preview_api


def register_router(app: FastAPI) -> None:
    """注册子路由；新模块在此 `include_router`。"""
    app.include_router(health_api)
    app.include_router(document_api)
    app.include_router(content_control_api)
    app.include_router(preview_api)
