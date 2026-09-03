import logging
import sys
from datetime import datetime
from typing import Dict

from word.config.path import LOG_DIR

class LoggerManager:
    """日志记录器管理器"""

    def __init__(
        self,
        logger_name: str,
        console_level: int = logging.WARNING,
        file_level: int = logging.INFO,
    ):
        """
        初始化日志记录器

        Args:
            logger_name: 日志记录器名称
            console_level: 控制台输出日志级别
            file_level: 文件输出日志级别
            log_dir: 日志文件存放目录
        """
        # 创建日志存放目录
        self.save_dir = LOG_DIR / logger_name
        self.save_dir.mkdir(parents=True, exist_ok=True)

        self.logger_name = logger_name
        self.console_level = console_level
        self.file_level = file_level

        # 创建日志记录器
        self.logger = logging.getLogger(logger_name)
        self.logger.setLevel(logging.DEBUG)  # 设置最低级别为DEBUG

        # 避免重复添加处理器
        if not self.logger.handlers:
            self._setup_handlers()

    def _setup_handlers(self):
        """设置日志处理器"""
        # 创建格式化器
        console_formatter = logging.Formatter(
            fmt="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        file_formatter = logging.Formatter(
            fmt="%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        # 设置控制台处理器
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(self.console_level)
        console_handler.setFormatter(console_formatter)

        # 设置文件处理器
        log_filename = f"{self.logger_name}_{datetime.now().strftime('%Y%m%d')}.log"
        log_filepath = self.save_dir / log_filename

        file_handler = logging.FileHandler(
            filename=log_filepath, mode="a", encoding="utf-8"
        )
        file_handler.setLevel(self.file_level)
        file_handler.setFormatter(file_formatter)

        # 添加处理器到日志记录器
        self.logger.addHandler(console_handler)
        self.logger.addHandler(file_handler)

    def get_logger(self) -> logging.Logger:
        """获取日志记录器实例"""
        return self.logger

    def debug(self, message: str):
        """记录DEBUG级别日志"""
        self.logger.debug(message)

    def info(self, message: str):
        """记录INFO级别日志"""
        self.logger.info(message)

    def warning(self, message: str):
        """记录WARNING级别日志"""
        self.logger.warning(message)

    def error(self, message: str):
        """记录ERROR级别日志"""
        self.logger.error(message)

    def critical(self, message: str):
        """记录CRITICAL级别日志"""
        self.logger.critical(message)


class LoggerFactory:
    def __init__(self):
        self.logger_mgr: Dict[str, LoggerManager] = {}

    def get_logger(self, logger_name: str) -> LoggerManager:
        """
        获取日志记录器实例，没有则创建
        """
        logger = self.logger_mgr.get(logger_name)
        if not logger:
            logger = LoggerManager(logger_name)
            self.logger_mgr[logger_name] = logger
        return logger


__LOGGER_FACTORY: LoggerFactory = LoggerFactory()


def get_logger(logger_name: str = "word") -> LoggerManager:
    return __LOGGER_FACTORY.get_logger(logger_name)
