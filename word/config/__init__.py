"""配置目录：路径、常量、运行时 settings。"""

from word.config.constants import (
    DEFAULT_HOST,
    DEFAULT_PORT,
    EASYWORD_API_VERSION,
    SERVICE_NAME,
)
from word.config.settings import get_host, get_port

__all__ = [
    "DEFAULT_HOST",
    "DEFAULT_PORT",
    "EASYWORD_API_VERSION",
    "SERVICE_NAME",
    "get_host",
    "get_port",
]
