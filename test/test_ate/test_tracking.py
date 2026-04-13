
import datetime
import os
import sys
from contextlib import redirect_stdout
from pathlib import Path
from test import working_path
from test.conftest import parametrize_from_json
from unittest.mock import patch

from xls_management.xlsx.workbook import Workbook


def fake_msgbox_no(msg:str)->bool:
    print(f'{msg}\nNo')
    sys.stdout.flush()
    return False

def fake_print(*vargs,**kvargs):
    print(*vargs,**kvargs)
    sys.stdout.flush()

def test_ATEStatus_perform_status():
    # ensure fresh imports so patched functions are picked up by modules
    old_stdout = sys.stdout
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]

    # prepare a list of file paths to be returned by the file picker
    file_path = working_path / "../ATEStatus_perfom_status.txt"
    # side_effect list long enough for repeated calls
    file_list = [
        working_path / '../in/012_MEB21/MEB21_Statistik_Testing.xlsx',
        working_path / '../in/012_MEB21/Alle Verifikationskriterien.xlsx',
        working_path / '../in/012_MEB21/Alle Absicherungsaufträge.xlsx',
        working_path / '../in/012_MEB21/Alle Testfälle.xlsx',
        working_path / '../in/012_MEB21/MasterFeatureplan.xlsx',
    #    working_path / '../in/trial_Master.xlsx',
    ]
    with file_path.open("w") as f:
        sys.std_out = f
        with(
            patch('xls_management.tui.file_picker.path_from_file_picker', side_effect=file_list),
            patch('xls_management.ate.project.project_combo_box', return_value=('MEB21', False)),
            patch('xls_management.tui.msgbox.msgbox', new=fake_print),
            patch('xls_management.tui.yes_no_form.yes_no_msgbox', new=fake_msgbox_no),
            patch('xls_management.ate.tracking.date') as mock_date,
        ):
            mock_date.today.return_value=datetime.date(2026, 4, 10)
            fake_print(f'....{__name__}')
            # import after patches so module-level imports pick up the patched functions
            from xls_management.shell.ate import ATEStatus

            ate_status = ATEStatus()
            #sheets = ate_status.output_workbook.sheet_names()
            #fake_print(', '.join(sheets))
            fake_print('..starting perform_status')
            ate_status.perform_status()
            fake_print('..perform_status ended')

            # project and flag were set by the mocked combo box
            assert ate_status.project == 'MEB21'
            assert ate_status.use_predecessor_ids is False
            
            #sheets_now = ate_status.output_workbook.sheet_names()
            #fake_print(','.join(sheets_now))
            #assert len(sheets_now) == len(sheets) + 1
        sys.stdout = old_stdout

def test_ATEStatus_perform_status_uc_config():
    """
    A repesentative sample of missing TD Status rows; comparing to VBA execution output
    is used.
    """
    # ensure fresh imports so patched functions are picked up by modules
    old_stdout = sys.stdout
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]

    # prepare a list of file paths to be returned by the file picker
    file_path = working_path / "../ATEStatus_perfom_status.txt"
    # side_effect list long enough for repeated calls
    output_path = working_path / '../out/012_MEB21/output_py.xlsx'
    if not output_path.parent.exists():
        os.makedirs(output_path.parent, exist_ok=True)
    with file_path.open("w") as f:
        sys.std_out = f
        with(
            patch('xls_management.ate.project.project_combo_box', return_value=('MEB21', False)),
            patch('xls_management.tui.msgbox.msgbox', new=fake_print),
            patch('xls_management.tui.yes_no_form.yes_no_msgbox', new=fake_msgbox_no),
            patch('xls_management.ate.tracking.date') as mock_date,
        ):
            mock_date.today.return_value=datetime.date(2026, 3, 23)
            fake_print(f'....{__name__}')
            # import after patches so module-level imports pick up the patched functions
            from xls_management.shell.ate import ATEStatus

            ate_status = ATEStatus()
            ate_status.config.config_from(working_path / '..\\in\\uc_config.yml')
            #sheets = ate_status.output_workbook.sheet_names()
            #fake_print(', '.join(sheets))
            fake_print('..starting perform_status')
            ate_status.perform_status(output_path)
            fake_print('..perform_status ended')

            # project and flag were set by the mocked combo box
            assert ate_status.project == 'MEB21'
            assert ate_status.use_predecessor_ids is False
            
            # check expected output
            w = Workbook(output_path)
            names = [name for name in w.sheet_names() if name[:2] == 'TD']
            assert len(names) == 1
            fake_print(','.join(names))
            with w.reader() as r:
                ws = w.sheet(names[0])
                assert len(ws) == 10
        sys.stdout = old_stdout

@parametrize_from_json("../test_data/in/status.json")
def test_ATEStatus_perform_status_req_id(file_list, output_path, project):
    """
    Testfalle in py should be the same in ATE_Status worksheet comparing to VBA execution output
    representative sample data is used.
    """
    # ensure fresh imports so patched functions are picked up by modules
    old_stdout = sys.stdout
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]
    # prepare a list of file paths to be returned by the file picker
    file_path = working_path / "../ATEStatus_perfom_status.txt"
    # side_effect list long enough for repeated calls
    output_path = working_path / output_path
    if not output_path.parent.exists():
        os.makedirs(output_path.parent, exist_ok=True)
    with file_path.open("w") as f:
        sys.std_out = f
        with(
            patch('xls_management.tui.file_picker.path_from_file_picker', side_effect=[working_path / p for p in file_list]),
            patch('xls_management.ate.project.project_combo_box', return_value=(project, False)),
            patch('xls_management.tui.msgbox.msgbox', new=fake_print),
            patch('xls_management.tui.yes_no_form.yes_no_msgbox', new=fake_msgbox_no),
            patch('xls_management.ate.tracking.date') as mock_date,
        ):
            mock_date.today.return_value=datetime.date(2026, 3, 23)
            fake_print(f'....{__name__}')
            # import after patches so module-level imports pick up the patched functions
            from xls_management.shell.ate import ATEStatus

            ate_status = ATEStatus()
            #sheets = ate_status.output_workbook.sheet_names()
            #fake_print(', '.join(sheets))
            fake_print('..starting perform_status')
            ate_status.perform_status(output_path)
            fake_print('..perform_status ended')

            # project and flag were set by the mocked combo box
            assert ate_status.project == project
            assert ate_status.use_predecessor_ids is False
            
            # check expected output
            w = Workbook(output_path)
            names = [name for name in w.sheet_names() if name[:2] == 'TD']
            assert len(names) == 1
            fake_print(','.join(names))
            with w.reader() as r:
                ws = w.sheet(names[0])
                #assert len(ws) == 10
        sys.stdout = old_stdout

def test_ATEStatus_config():
    from xls_management.shell.ate import ATEStatus

    ate_status = ATEStatus()
    assert ate_status.config is not None
    file_path_BsM:str|None = ate_status.config.get('workbook_path_BsM')
    assert file_path_BsM is not None

def test_ATEStatus_read_blacklist():
    from xls_management.shell.ate import ATEStatus

    ate_status:ATEStatus = ATEStatus()
    ate_status.read_blacklist_LAHB()
    assert ate_status.blacklist_LAHB is not None
    assert isinstance(ate_status.blacklist_LAHB, tuple) is True
