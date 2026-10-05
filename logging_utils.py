import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler


def configure_app_logger(base_dir, name="erp_beta"):
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.propagate = False
    if not logger.handlers:
        handler = TimedRotatingFileHandler(
            os.path.join(log_dir, "erp.log"),
            when="midnight",
            interval=1,
            backupCount=30,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
        ))
        logger.addHandler(handler)
    return logger


def install_global_exception_logging(logger):
    original = sys.excepthook

    def hook(exc_type, exc_value, exc_tb):
        try:
            logger.exception("Excepcion no controlada", exc_info=(exc_type, exc_value, exc_tb))
        finally:
            original(exc_type, exc_value, exc_tb)

    sys.excepthook = hook
    return hook
