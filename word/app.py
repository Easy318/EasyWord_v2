"""FastAPI 应用工厂。"""

from __future__ import annotations

from contextlib import asynccontextmanager
from typing import AsyncIterator

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from word.config.constants import SERVICE_NAME
from word.logger import get_logger
from word.router import register_router
from word.runtime.com_host import com_host
from word.runtime.document_binding import document_binding
from word.runtime.errors import EasyWordError
from word.runtime.word_app import word_app

_log = get_logger("app")


@asynccontextmanager
async def lifespan(_app: FastAPI) -> AsyncIterator[None]:
    com_host.start()
    _log.info("EasyWord lifespan startup")
    try:
        yield
    finally:
        _log.info("EasyWord lifespan shutdown")
        try:
            if com_host.started:
                com_host.submit(lambda: document_binding.close(save=True))
                com_host.submit(word_app.shutdown)
        except Exception as exc:  # noqa: BLE001
            _log.warning(f"shutdown cleanup: {exc}")
        com_host.stop()


def create_app() -> FastAPI:
    app = FastAPI(title=SERVICE_NAME, docs_url="/docs", redoc_url=None, lifespan=lifespan)
    # 渲染进程来自 Vite/Electron 源站，对本机 sidecar 的 POST 会先发 OPTIONS 预检
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=False,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    register_router(app)

    @app.exception_handler(EasyWordError)
    async def easyword_error_handler(_request: Request, exc: EasyWordError) -> JSONResponse:
        return JSONResponse(
            status_code=exc.status_code,
            content={"code": exc.code, "message": exc.message},
        )

    return app


app = create_app()
