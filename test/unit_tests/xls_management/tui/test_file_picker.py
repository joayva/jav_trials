from pathlib import Path

import pytest
from textual_fspicker import FileOpen, Filters

from xls_management.tui.file_picker import FilePickerApp


def test_file_picker_initialization():
    picker = FilePickerApp(
        path = Path("."), 
        filters = Filters(
            ("Excel", lambda p: p.suffix.lower() == ".xlsx"),
            ("CSV", lambda p: p.suffix.lower() == ".csv"),
            ("All", lambda _: True),
        ),
        title ="Open",
    )
    assert picker.path == Path('.')

@pytest.mark.xfail(reason="Test should be implemented")
@pytest.mark.asyncio
async def test_file_picker_select_file():
    raise NotImplementedError
    picker = FilePickerApp(
        path = Path("."), 
        filters = Filters(
            ("Excel", lambda p: p.suffix.lower() == ".xlsx"),
            ("CSV", lambda p: p.suffix.lower() == ".csv"),
            ("All", lambda _: True),
        ),
        title ="Open",
    )
    async with picker.run_test() as pilot:
        pilot.post_message(
            FileOpen(
                sender=pilot,
                path=Path('test.xlsx'),
            )
        )
        pilot.pause(2.0)
        assert picker.selected == 'test.xlsx'

@pytest.mark.xfail(reason="Test should be implemented")
def test_file_picker_edge_case_invalid_file():
    raise NotImplementedError
    picker = FilePickerApp()
    with pytest.raises(FileNotFoundError):
        picker.select_file("nonexistent.xlsx")

@pytest.mark.xfail(reason="Test should be implemented")
def test_file_picker_integration_with_main():
    raise NotImplementedError
    from xls_management.tui.main import MainApp
    picker = FilePickerApp()
    app = MainApp()
    app.use_picker(picker)
    assert app.picker is picker
