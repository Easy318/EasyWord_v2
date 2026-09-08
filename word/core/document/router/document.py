"""文档会话路由。"""

from __future__ import annotations

from fastapi import APIRouter

from word.core.document import service
from word.core.document.schema import (
    DocumentCloseIn,
    DocumentFinalizeIn,
    DocumentOpenIn,
    DocumentOpenOut,
    DocumentPersistIn,
    DocumentStatusOut,
    OkOut,
)

document_api = APIRouter(prefix="/document", tags=["document"])


@document_api.post("/open", response_model=DocumentOpenOut)
def open_document(body: DocumentOpenIn) -> DocumentOpenOut:
    result = service.open_document(body.projectId, body.templatePath)
    return DocumentOpenOut.model_validate(result)


@document_api.post("/save", response_model=OkOut)
def save_document() -> OkOut:
    service.save_document()
    return OkOut()


@document_api.post("/persist", response_model=OkOut)
def persist_document(body: DocumentPersistIn) -> OkOut:
    """按路径落盘且保持打开（批量生成前同步；不依赖绑定状态）。"""
    service.persist_document(body.templatePath)
    return OkOut()


@document_api.post("/close", response_model=OkOut)
def close_document(body: DocumentCloseIn | None = None) -> OkOut:
    save = True if body is None else body.save
    service.close_document(save=save)
    return OkOut()


@document_api.post("/finalize", response_model=OkOut)
def finalize_document(body: DocumentFinalizeIn) -> OkOut:
    """按路径保存并关闭模板（退出项目时调用，不依赖绑定状态）。"""
    service.finalize_document(body.templatePath)
    return OkOut()


@document_api.get("/status", response_model=DocumentStatusOut)
def document_status() -> DocumentStatusOut:
    return DocumentStatusOut.model_validate(service.document_status())
