import logging
import sys
import traceback
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Any

# Force UTF-8 for stdout/stderr to avoid GBK encoding errors on Chinese Windows
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

LOG_PATH = 'logs/out.log'


# ================= 终端 ANSI 颜色配置 =================
class Colors:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    BLACK = "\033[30m"
    RED = "\033[31m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    BLUE = "\033[34m"
    MAGENTA = "\033[35m"
    CYAN = "\033[36m"
    WHITE = "\033[37m"
    GRAY = "\033[90m"


LEVEL_COLORS = {
    logging.DEBUG: Colors.CYAN,
    logging.INFO: Colors.BLUE,
    logging.WARNING: Colors.YELLOW,
    logging.ERROR: Colors.RED,
    logging.CRITICAL: Colors.MAGENTA + Colors.BOLD,
}


# ================= 格式化器 =================
class ColoredStructuredFormatter(logging.Formatter):
    """用于控制台输出的彩色格式化器"""

    def format(self, record):
        level_color = LEVEL_COLORS.get(record.levelno, Colors.WHITE)
        time_str = datetime.fromtimestamp(record.created).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        colored_time = f"{Colors.GRAY}{time_str}{Colors.RESET}"

        message = record.getMessage()
        kv_part = ""
        if " | " in message:
            msg_text, kv_text = message.split(" | ", 1)
            kv_part = f" {Colors.CYAN}({kv_text}){Colors.RESET}"
            message = msg_text

        colored_level_name = f"{level_color}[{record.levelname}: {record.name}]{Colors.RESET}"
        gray_info = f"{Colors.GRAY}{record.filename}:{record.lineno} | {record.funcName}{Colors.RESET}"

        return f"{colored_time} {colored_level_name} {message}{kv_part} {gray_info}"


class FileStructuredFormatter(logging.Formatter):
    """用于文件输出的纯净格式化器 (无 ANSI 颜色码)"""

    def __init__(self):
        fmt = "%(asctime)s.%(msecs)03d [%(levelname)s: %(name)s] %(message)s (%(filename)s:%(lineno)d | %(funcName)s)"
        datefmt = "%Y-%m-%d %H:%M:%S"
        super().__init__(fmt=fmt, datefmt=datefmt)

    def format(self, record):
        message = record.getMessage()
        if " | " in message:
            msg_text, kv_text = message.split(" | ", 1)
            message = f"{msg_text} ({kv_text})"
        record.msg = message
        return super().format(record)


# ================= 核心封装类 =================
class Logger:
    """
    线程安全的日志记录器封装。
    使用组合而非继承，避免了猴子补丁的侵入性。
    """

    def __init__(self, name: str = None, level: int = logging.INFO, log_file: str = LOG_PATH,
                 max_bytes: int = 5 * 1024 * 1024, backup_count: int = 5):
        self._logger = logging.getLogger(name)

        # 防止重复添加 Handler
        if not self._logger.handlers:
            self._logger.setLevel(level)

            # 1. 控制台 Handler
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(ColoredStructuredFormatter())
            self._logger.addHandler(console_handler)

            # 2. 文件 Handler
            if log_file:
                file_handler = RotatingFileHandler(
                    filename=log_file,
                    maxBytes=max_bytes,
                    backupCount=backup_count,
                    encoding='utf-8'
                )
                file_handler.setFormatter(FileStructuredFormatter())
                self._logger.addHandler(file_handler)

            self._logger.propagate = False

    def debug(self, msg: str, **kwargs: Any) -> None:
        self._log(self._logger.debug, msg, **kwargs)

    def info(self, msg: str, **kwargs: Any) -> None:
        self._log(self._logger.info, msg, **kwargs)

    def warning(self, msg: str, **kwargs: Any) -> None:
        self._log(self._logger.warning, msg, **kwargs)

    def error(self, msg: str, **kwargs: Any) -> None:
        tb_str = traceback.format_exc()
        kwargs['trace'] = tb_str
        self._log(self._logger.error, msg, **kwargs)

    def critical(self, msg: str, **kwargs: Any) -> None:
        self._log(self._logger.critical, msg, **kwargs)

    def exception(self, msg: str, **kwargs: Any) -> None:
        tb_str = traceback.format_exc()
        kwargs['trace'] = tb_str
        self._log(self._logger.exception, msg, **kwargs)

    def _log(self, log_func, msg: str, **kwargs: Any) -> None:
        """内部统一处理结构化参数"""
        if kwargs:
            kv_str = ", ".join(f"{k}={v}" for k, v in kwargs.items())
            msg = f"{msg} | {kv_str}"
        log_func(msg, stacklevel=3)


# ================= 全局获取函数 =================
def get_logger(name: str = None, level: int = logging.DEBUG, log_file: str = LOG_PATH, max_bytes: int = 5 * 1024 * 1024,
               backup_count: int = 5) -> Logger:
    """
    获取封装后的 Logger 实例。
    外部调用者无需知道标准库的细节。
    """
    return Logger(name, level, log_file, max_bytes, backup_count)
