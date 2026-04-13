from pathlib import Path
from test import working_path
from unittest.mock import patch

from xls_management.ate.om.db_info import DBInfo
from xls_management.utils.tools import all_in_sequence


def test_db_info_init():
    data:DBInfo = DBInfo(attributes=('License plate', 'Brand', 'Modell'))
    assert all_in_sequence(('License plate', 'Brand', 'Modell'),data.attributes)
    assert len(data.attributes) == 3

def test_db_info_read_file():
    import sys
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]
    file_path: Path = working_path / f"test/data/example01.xlsx"
    with patch('xls_management.tui.file_picker.path_from_file_picker', return_value=file_path):
        from xls_management.ate.om.db_info import DBInfo
        from xls_management.utils.tools import all_in_sequence
        data = DBInfo(attributes=('License plate', 'Brand', 'Modell'))
        success = data.read_file("Test")
        assert success is True
        assert data.error_msg == ""
        assert all_in_sequence(data.attributes, data.columns.keys())
        assert data.sheet_name == "Cars"
        assert len(data.columns["License plate"]) == 3
        assert data.columns['License plate'][0] == 'DE2456HBZ'
        assert data.columns['Brand'][2] == 'Audi'
        assert data.columns['Modell'][1] == 'Polo'

def test_db_info_read_file_na():
    import sys
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]
    file_path: Path = working_path / f"test/data/example_na.xlsx"
    with patch('xls_management.tui.file_picker.path_from_file_picker', return_value=file_path):
        from xls_management.ate.om.db_info import DBInfo
        from xls_management.utils.tools import all_in_sequence
        data = DBInfo(attributes=('License plate', 'Brand', 'Modell'))
        success = data.read_file("Test")
        assert success is True
        assert data.error_msg == ""
        assert all_in_sequence(data.attributes, data.columns.keys())
        assert data.sheet_name == "Cars"
        assert len(data.columns["License plate"]) == 3
        assert data.columns['License plate'][0] == 'DE2456HBZ'
        assert data.columns['Brand'][2] == 'Audi'
        assert data.columns['Modell'][1] == 'n/a'

def test_db_info_read_file_error():
    import sys
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]
    file_path: Path = working_path / f"test/data/example01.xlsx"
    with patch('xls_management.tui.file_picker.path_from_file_picker', return_value=file_path):
        from xls_management.ate.om.db_info import DBInfo
        from xls_management.utils.tools import all_in_sequence
        data = DBInfo(attributes=('Name', 'License plate', 'Brand', 'Modell'))
        success = data.read_file("Test")
        assert success is False
        assert data.error_msg != ""
        lines = data.error_msg.splitlines()
        assert "People" in lines[0]
        assert "Name" in lines[1]
        for attribute in data.attributes[1:]:
            assert attribute in lines[0]
            assert attribute not in lines[1]
        assert "Cars" in lines[1]
        assert "Name" in lines[1]

            