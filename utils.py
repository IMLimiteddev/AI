# utils.py
from datetime import datetime
import logging

def format_date_dmy(date_str, input_format="%d.%m.%Y", output_format="%d.%m.%Y"):
    """Formats a date string, returns original or None on error."""
    if not date_str:
        return None
    try:
        # Handle different input formats if necessary
        dt_obj = datetime.strptime(date_str, input_format)
        return dt_obj.strftime(output_format)
    except ValueError:
        logging.warning(f"Could not parse date '{date_str}' with format '{input_format}'")
        # Attempt other common formats? Or just return original/None
        return date_str # Return original string if parsing fails