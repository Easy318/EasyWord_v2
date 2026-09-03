"""RenderContext 镜像（与 QHub 线格式一致，禁止 import QHub）。"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field, model_validator

from word.core.preview.schema.common import (
    CAMEL_CONFIG,
    JINJA_RESERVED_NAMES,
    SCHEMA_VERSION,
    RenderSource,
    RenderStatus,
    ValueDtype,
)


class RenderValue(BaseModel):
    value: Any
    dtype: ValueDtype
    shape: list[int] | None = None
    source: RenderSource
    status: RenderStatus = "active"

    model_config = CAMEL_CONFIG


class BurstBinding(BaseModel):
    var_name: str = Field(..., alias="varName")
    value: Any

    model_config = CAMEL_CONFIG


class RenderContext(BaseModel):
    schema_version: str = Field(SCHEMA_VERSION, alias="schemaVersion")
    project_id: str = Field(..., alias="projectId")
    burst: BurstBinding | None = Field(
        None,
        description="未配置展开变量时为 null，此时不注入 Jinja 保留名",
    )
    entries: dict[str, RenderValue] = Field(default_factory=dict)

    model_config = CAMEL_CONFIG

    @model_validator(mode="after")
    def _no_reserved_entry_keys(self) -> RenderContext:
        clash = JINJA_RESERVED_NAMES.intersection(self.entries)
        if clash:
            raise ValueError(f"entries 键不可使用 Jinja 保留名: {sorted(clash)}")
        return self
