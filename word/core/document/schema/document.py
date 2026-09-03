"""文档会话 schema。"""

from __future__ import annotations

from pydantic import BaseModel, Field


class DocumentOpenIn(BaseModel):
    projectId: str = Field(min_length=1, description="项目 ID")
    templatePath: str = Field(min_length=1, description="本机模板绝对路径")


class DocumentOpenOut(BaseModel):
    projectId: str
    templatePath: str
    rebound: bool = Field(description="是否重绑到已打开的同路径文档")


class DocumentCloseIn(BaseModel):
    save: bool = Field(default=True, description="关闭前是否保存")


class DocumentFinalizeIn(BaseModel):
    templatePath: str = Field(min_length=1, description="本机模板绝对路径")


class DocumentStatusOut(BaseModel):
    bound: bool
    projectId: str | None = None
    templatePath: str | None = None
    wordAvailable: bool
    documentOpen: bool


class OkOut(BaseModel):
    ok: bool = True
