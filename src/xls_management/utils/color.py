from enum import Enum


# ANSI escape codes for colors
class Color(str,Enum):
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    MAGENTA = "\033[95m"
    CYAN = "\033[96m"

RESET = "\033[0m"  # Reset to default color

def ansi_color(text:str, color:Color) -> str:
    return f"{color.value}{text}{RESET}"