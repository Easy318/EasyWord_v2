"""文档会话服务（在 COM 线程执行）。"""

from __future__ import annotations

from word.runtime.com_host import com_host
from word.runtime.document_binding import document_binding


def open_document(project_id: str, template_path: str) -> dict[str, object]:
    return com_host.submit(lambda: document_binding.open(project_id, template_path))


def save_document() -> None:
    com_host.submit(document_binding.save)


def close_document(save: bool = True) -> None:
    com_host.submit(lambda: document_binding.close(save=save))


def finalize_document(template_path: str) -> None:
    com_host.submit(lambda: document_binding.finalize_by_path(template_path))


def document_status() -> dict[str, object]:
    return com_host.submit(document_binding.status)
