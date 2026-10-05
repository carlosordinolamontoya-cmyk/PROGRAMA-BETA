"""Nucleo reutilizable del ERP Beta.

V44 inicia la separacion progresiva del archivo monolitico sin cambiar la
experiencia del usuario: la aplicacion sigue abriendose desde app_kardex.py.
"""

from .money import (
    D,
    q_qty,
    q_unit,
    q_tc,
    q_money,
    calc_price_values,
    money_sum,
    money_fmt,
)
from .dates import iso_now, normalize_datetime_text
from .correlatives import AnnualCorrelativeService
from .logging_utils import configure_app_logger, install_global_exception_logging
from .sqlite_utils import apply_sqlite_runtime_pragmas, ensure_performance_indexes

__all__ = [
    "D", "q_qty", "q_unit", "q_tc", "q_money", "calc_price_values",
    "money_sum", "money_fmt", "iso_now", "normalize_datetime_text",
    "AnnualCorrelativeService", "configure_app_logger",
    "install_global_exception_logging", "apply_sqlite_runtime_pragmas",
    "ensure_performance_indexes",
]
