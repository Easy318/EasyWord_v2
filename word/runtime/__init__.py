"""Word COM 运行时：STA host、Application 单例、文档绑定。"""

from __future__ import annotations

from word.runtime.com_host import com_host
from word.runtime.document_binding import document_binding
from word.runtime.word_app import word_app

__all__ = ["com_host", "document_binding", "word_app"]
