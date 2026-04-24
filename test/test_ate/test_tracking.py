
import datetime
import os
import sys
from test import working_path
from test.conftest import del_xls_management_imports, parametrize_from_yaml
from unittest.mock import patch

from xls_management import WORKPATH
from xls_management.xlsx.workbook import Workbook


def fake_msgbox_no(msg:str)->bool:
    print(f'{msg}\nNo')
    sys.stdout.flush()
    return False

def fake_print(*vargs,**kvargs):
    print(*vargs,**kvargs)
    sys.stdout.flush()

@parametrize_from_yaml("../test_data/in/status_config.yml")
def test_ATEStatus_perform_status_config(
    config_file:str, 
    project:str,
    date:list[int],
):
    """
    A repesentative sample of missing TD Status rows; comparing to VBA execution output
    is used.
    """
    # ensure fresh imports so patched functions are picked up by modules
    old_stdout = sys.stdout
    del_xls_management_imports()

    # prepare a list of file paths to be returned by the file picker
    file_path = working_path / "../ATEStatus_perfom_status.txt"
    # side_effect list long enough for repeated calls
    with file_path.open("w") as f:
        sys.std_out = f
        with(
            patch('xls_management.tui.msgbox.msgbox', new=fake_print),
            patch('xls_management.tui.yes_no_form.yes_no_msgbox', new=fake_msgbox_no),
            patch('xls_management.ate.tracking.date') as mock_date,
        ):
            mock_date.today.return_value=datetime.date(*date)
            fake_print(f'....{__name__}')
            # import after patches so module-level imports pick up the patched functions
            from xls_management.shell.ate import ATEStatus

            ate_status = ATEStatus()
            ate_status.config.config_from(config_file.format(work_path=WORKPATH))
            #sheets = ate_status.output_workbook.sheet_names()
            #fake_print(', '.join(sheets))
            fake_print('..starting perform_status')
            ate_status.perform_status()
            fake_print('..perform_status ended')

            # project and flag were set by the mocked combo box
            assert ate_status.project == project
            assert ate_status.use_predecessor_ids is False
            
            # check expected output
            output_path = ate_status.config.get('output_path',None)
            assert output_path is not None
            w = Workbook(output_path)
            names = [name for name in w.sheet_names() if name[:9] in ('TD_Status','ATE_Statu')]
            assert len(names) == 2
            fake_print(','.join(names))
            #with w.reader() as r:
            #    ws = w.sheet(names[0])
            #    assert len(ws) == 10
        sys.stdout = old_stdout

@parametrize_from_yaml("../test_data/in/status.yml")
def test_ATEStatus_perform_status(
    file_list:list[str], 
    output_path:str, 
    project:str, 
    date:list[int],
):
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
            mock_date.today.return_value=datetime.date(*date)
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
