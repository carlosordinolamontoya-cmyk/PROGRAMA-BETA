from datetime import datetime

ISO_DATETIME_FMT = "%Y-%m-%d %H:%M:%S"
LEGACY_DATETIME_FMT = "%d/%m/%Y %H:%M:%S"


def iso_now():
    return datetime.now().strftime(ISO_DATETIME_FMT)


def normalize_datetime_text(value):
    """Normaliza timestamps internos a YYYY-MM-DD HH:MM:SS.

    Fechas de negocio (solo fecha) no deben enviarse a esta funcion.
    """
    text = str(value or "").strip()
    if not text:
        return text
    for fmt in (ISO_DATETIME_FMT, LEGACY_DATETIME_FMT):
        try:
            return datetime.strptime(text, fmt).strftime(ISO_DATETIME_FMT)
        except ValueError:
            pass
    # Compatibilidad con valores que incluyen microsegundos.
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%d/%m/%Y %H:%M:%S.%f"):
        try:
            return datetime.strptime(text, fmt).strftime(ISO_DATETIME_FMT)
        except ValueError:
            pass
    return text
