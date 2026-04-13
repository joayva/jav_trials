import os
import test
from pathlib import Path
from test import working_path

from xls_management import WORKPATH


def test_config_create(monkeypatch):
    clean("test/data/config.yml")
    monkeypatch.setattr("xls_management.ROOTPATH",working_path / "test/data")
    from xls_management.config import ATEConfig
    ate = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        WORKPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsx',
    )

def test_config_open(monkeypatch, tmp_path):
    target_path = tmp_path / "config.yml"
    assert target_path.exists() is False
    monkeypatch.setattr("xls_management.ROOTPATH", tmp_path)
    from xls_management.config import ATEConfig
    ate:ATEConfig = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        WORKPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsx',
    )
    target:str|None = ate.get('workbook_path_BsM')
    assert target is not None
    del(ate)
    assert (tmp_path / "config.yml").exists() is True
    ate = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        WORKPATH /
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