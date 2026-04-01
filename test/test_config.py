import os
import test
from test import working_path
from xls_management import HOMEPATH
from pathlib import Path

def test_config_create(monkeypatch):
    clean("test/data/config.yml")
    monkeypatch.setattr("xls_management.ROOTPATH",working_path / "test/data")
    from xls_management.config import ATEConfig
    ate = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        HOMEPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsx',
    )

def test_config_open(monkeypatch):
    clean("test/data/config.yml")
    assert (working_path / "test/data/config.yml").exists() is False
    monkeypatch.setattr("xls_management.ROOTPATH",working_path / "test/data")
    from xls_management.config import ATEConfig
    ate:ATEConfig = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        HOMEPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsx',
    )
    target:str|None = ate.get('workbook_path_BsM')
    assert target is not None
    del(ate)
    assert (working_path / "test/data/config.yml").exists() is True
    ate = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        HOMEPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsx',
    )
    target:str|None = ate.get('workbook_path_BsM')
    assert target is not None

def test_config_worksheet_widths(monkeypatch):
    #monkeypatch.setattr("xls_management.ROOTPATH",working_path / "test/data")
    from xls_management.config import ATEConfig
    ate = ATEConfig()
    assert 'worksheet_widths' in ate.config.keys()
    ww = ate.config['worksheet_widths']
    assert 'ATE_Status' in ww.keys()
    assert 'TD_Status' in ww.keys()

def clean(working_file):
    file_path = working_path / working_file
    if file_path.exists():
        os.remove(file_path)