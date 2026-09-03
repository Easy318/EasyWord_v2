"""EasyWord 业务错误（映射为 HTTP JSON）。"""

from __future__ import annotations


class EasyWordError(Exception):
    def __init__(self, code: str, message: str, status_code: int = 400) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.status_code = status_code
