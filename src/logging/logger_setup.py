# -*- coding: utf-8 -*-
import logging
import logging.handlers
from pathlib import Path

# Cố gắng import cấu hình từ src.__init__, có fallback nếu thiếu
try:
    from .. import APP_NAME as _APP_NAME
except Exception:
    _APP_NAME = "word_to_pdf"

try:
    from .. import LOG_LEVEL as _LOG_LEVEL
except Exception:
    _LOG_LEVEL = "INFO"

try:
    from .. import FILE_LOGGING as _FILE_LOGGING
except Exception:
    _FILE_LOGGING = False

try:
    from .. import MAX_FILE_SIZE as _MAX_FILE_SIZE
except Exception:
    _MAX_FILE_SIZE = 1_000_000

try:
    from .. import BACKUP_COUNT as _BACKUP_COUNT
except Exception:
    _BACKUP_COUNT = 3

try:
    from .. import LOG_PATH as _LOG_PATH
except Exception:
    _LOG_PATH = Path.cwd() / "logs"


def setup_logger(
    name: str = None,
    level: str = None,
    log_level: str = None,
    log_format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    log_to_file: bool = None,
    log_dir: Path = None
) -> logging.Logger:
    """Khởi tạo logger theo cấu hình."""
    # Ưu tiên tham số level => log_level => mặc định _LOG_LEVEL
    if level is not None:
        log_level = level
    if log_level is None:
        log_level = _LOG_LEVEL

    if name is None:
        name = _APP_NAME.lower().replace(' ', '_')

    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))

    # Tránh add handler trùng lặp nếu đã khởi tạo
    if logger.handlers:
        return logger

    # Console handler
    formatter = logging.Formatter(log_format, datefmt="%Y-%m-%d %H:%M:%S")
    ch = logging.StreamHandler()
    ch.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    # File handler nếu bật
    if log_to_file is None:
        log_to_file = _FILE_LOGGING

    if log_to_file:
        if log_dir is None:
            log_dir = Path(_LOG_PATH)
        try:
            log_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        fh = logging.handlers.RotatingFileHandler(
            filename=str(log_dir / f"{name}.log"),
            maxBytes=int(_MAX_FILE_SIZE),
            backupCount=int(_BACKUP_COUNT),
            encoding="utf-8"
        )
        fh.setLevel(getattr(logging, log_level.upper(), logging.INFO))
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    return logger


def get_logger(name: str = None) -> logging.Logger:
    """Lấy logger với tên chuẩn hóa theo APP_NAME (hoặc __name__)."""
    if name is None:
        name = _APP_NAME.lower().replace(' ', '_')
    return logging.getLogger(name)
