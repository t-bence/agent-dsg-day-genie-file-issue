from datetime import datetime

from agents import function_tool


@function_tool
def get_current_time() -> str:
    """Get the current date and time."""
    return datetime.now().isoformat()
