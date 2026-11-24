"""Small value helpers shared by the readers and writers."""
import datetime as dt


def text(v) -> str:
    """Cell value as a trimmed string ('' for empty). Whole numbers lose the '.0'."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%Y-%m-%d')
    return str(v).strip()


def key(v) -> str:
    """Lookup key for order numbers / ASINs: text without inner spaces."""
    return text(v).replace(' ', '')


def to_int(v, default=0) -> int:
    try:
        return int(float(text(v)))
    except ValueError:
        return default
