import pandas as pd
import unicodedata
import re

def format_boolean(value):
    """Convert various boolean representations to yes/no."""
    if pd.isna(value):
        return None
    value = str(value).lower().strip()
    if value in ['enable', 'enabled', 'yes', 'true', '1']:
        return "yes"
    elif value in ['disable', 'disabled', 'no', 'false', '0']:
        return "no"
    return None

def fix_capacity(capacity_value):
    """Check if capacity is <= 4.000GB and change it to 10.000GB if true"""
    try:
        str_value = str(capacity_value).strip().upper()

        import re
        match = re.match(r"([0-9.]+)\s*(KB|MB|GB|TB)", str_value)
        if not match:
            return capacity_value

        value, unit = match.groups()
        value = float(value)

        if unit == "GB":
            if value < 4.000:
                return "10.000GB"
        elif unit in ("KB", "MB"):
            return "10.000GB"

    except (ValueError, TypeError):
        pass

    return capacity_value

def fix_description(description_value):
    """Sanitize the description string for command use"""
    if not isinstance(description_value, str):
        return description_value

    fixed = description_value.replace(' ', '_')

    fixed = unicodedata.normalize('NFKD', fixed).encode('ASCII', 'ignore').decode()

    fixed = re.sub(r'[^a-zA-Z0-9_\-@]', '', fixed)

    return fixed