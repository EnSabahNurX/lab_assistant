from datetime import datetime


def clean_value(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    value = str(value).strip()
    return value if value and value.lower() != "none" else None


def parse_date(date_value):
    if not date_value:
        return "0000-00-00"
    try:
        if isinstance(date_value, datetime):
            return date_value.strftime("%Y-%m-%d")
        if isinstance(date_value, str):
            for fmt in [
                "%Y-%m-%d",
                "%d.%m.%Y",
                "%m/%d/%Y",
                "%Y/%m/%d",
            ]:
                try:
                    return datetime.strptime(date_value, fmt).strftime("%Y-%m-%d")
                except ValueError:
                    continue
        return "0000-00-00"
    except Exception:
        return "0000-00-00"
