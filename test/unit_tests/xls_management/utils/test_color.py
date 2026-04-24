from test.conftest import parametrize_from_yaml

from xls_management import WORKPATH
from xls_management.utils.color import RESET, Color, ansi_color


def test_color():
    assert Color.RED == "\x1b[91m"
    assert Color.GREEN == "\x1b[92m"
    assert Color.YELLOW == "\x1b[93m"
    assert Color.BLUE == "\x1b[94m"
    assert Color.MAGENTA == "\x1b[95m"
    assert Color.CYAN == "\x1b[96m"

    assert RESET == "\x1b[0m"

@parametrize_from_yaml(file_path=f'{WORKPATH}\\vw_dev\\test_data\\in\\color.yml')
def test_ansi_color(text:str, color_key:str, expected_result:str) -> None:
    assert ansi_color(text, Color[color_key]) == expected_result