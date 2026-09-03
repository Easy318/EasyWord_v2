"""健康检查路由。"""

from __future__ import annotations

from fastapi import APIRouter

from word import __version__
from word.config.constants import EASYWORD_API_VERSION, SERVICE_NAME
from word.core.health.schema import HealthOut

health_api = APIRouter(tags=["health"])


@health_api.get("/health", response_model=HealthOut)
def health() -> HealthOut:
    return HealthOut(
        ok=True,
        service=SERVICE_NAME,
        version=__version__,
        apiVersion=EASYWORD_API_VERSION,
    )
