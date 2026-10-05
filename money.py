from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

Q_QTY = Decimal("0.001")
Q_UNIT = Decimal("0.0001")
Q_TC = Decimal("0.0001")
Q_MONEY = Decimal("0.01")
IGV = Decimal("0.18")
ONE = Decimal("1")


def D(value, default="0"):
    """Convierte a Decimal sin pasar por binario float cuando es posible."""
    if value is None or value == "":
        value = default
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value).strip().replace(",", "."))
    except (InvalidOperation, ValueError, TypeError):
        raise ValueError(f"Valor numerico invalido: {value}")


def _q(value, quantum):
    return D(value).quantize(quantum, rounding=ROUND_HALF_UP)


def q_qty(value):
    return _q(value, Q_QTY)


def q_unit(value):
    return _q(value, Q_UNIT)


def q_tc(value):
    return _q(value, Q_TC)


def q_money(value):
    return _q(value, Q_MONEY)


def calc_price_values(cantidad, precio_unitario, precio_incluye_igv=False):
    cantidad_d = q_qty(cantidad)
    precio_d = q_unit(precio_unitario)
    if cantidad_d <= 0:
        raise ValueError("La cantidad debe ser mayor a cero.")
    if precio_d <= 0:
        raise ValueError("El precio debe ser mayor a cero.")

    if precio_incluye_igv:
        precio_con = precio_d
        precio_sin = q_unit(precio_con / (ONE + IGV))
    else:
        precio_sin = precio_d
        precio_con = q_unit(precio_sin * (ONE + IGV))

    subtotal = q_money(cantidad_d * precio_sin)
    igv = q_money(subtotal * IGV)
    total = q_money(subtotal + igv)
    return precio_sin, precio_con, subtotal, igv, total


def money_sum(values):
    total = Decimal("0")
    for value in values:
        total += D(value)
    return q_money(total)


def money_fmt(value):
    return f"{q_money(value):,.2f}"
