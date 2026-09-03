"""健康检查响应契约。"""

from __future__ import annotations

from pydantic import BaseModel, Field


class HealthOut(BaseModel):
    ok: bool = Field(description="服务是否可用")
    service: str = Field(description="服务名")
    version: str = Field(description="服务版本")
    apiVersion: int = Field(description="HTTP 契约版本，与客户端 EXPECTED_API_VERSION 对齐")
