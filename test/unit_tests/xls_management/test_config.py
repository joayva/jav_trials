import os
import test
from pathlib import Path
from test import working_path
from test.conftest import del_xls_management_imports, parametrize_from_yaml

import pytest

from xls_management import WORKPATH as MYDATA


def test_config_create(monkeypatch, tmp_path):
    del_xls_management_imports()
    monkeypatch.setattr('xls_management.ROOTPATH',tmp_path / 'root_path')
    monkeypatch.setattr('xls_management.WORKPATH',tmp_path / 'work_path')
    from xls_management import WORKPATH
    from xls_management.config import ATEConfig
    ate = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        WORKPATH /
        'vw_dev/data/ATE-Status_Berichtsversion.xlsx',
    )

def test_config_open(monkeypatch, tmp_path):
    del_xls_management_imports()
    target_path = tmp_path / '.xls\config.yml'
    assert target_path.exists() is False
    monkeypatch.setattr('xls_management.WORKPATH', tmp_path)
    from xls_management import WORKPATH
    from xls_management.config import ATEConfig
    ate:ATEConfig = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        WORKPATH /
        'vw_dev/data/ATE-Status_Berichtsversion.xlsx',
    )
    target:str|None = ate.get('workbook_path_BsM')
    assert target is not None
    del(ate)
    assert target_path.exists() is True
    ate = ATEConfig()
    assert 'workbook_path_BsM' in ate.config.keys()
    assert ate.config['workbook_path_BsM'] == str(
        WORKPATH /
        'vw_dev/data/ATE-Status_Berichtsversion.xlsx',
    )
    target:str|None = ate.get('workbook_path_BsM')
    assert target is not None

def test_config_worksheet_widths(monkeypatch, tmp_path):
    del_xls_management_imports()
    monkeypatch.setattr('xls_management.ROOTPATH', tmp_path / 'root_path')
    monkeypatch.setattr('xls_management.WORKPATH', tmp_path / 'work_path')
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

def test_config_initialization(monkeypatch, tmp_path):
    del_xls_management_imports()
    monkeypatch.setattr('xls_management.ROOTPATH', tmp_path / 'root_path')
    monkeypatch.setattr('xls_management.WORKPATH', tmp_path / 'work_path')
    from xls_management.config import ATEConfig
    config = ATEConfig()
    assert hasattr(config, 'config_file')
    assert hasattr(config, 'config')
    assert config.config_file == tmp_path / 'work_path\.xls\config.yml'

def test_config_properties(monkeypatch, tmp_path):
    del_xls_management_imports()
    monkeypatch.setattr('xls_management.ROOTPATH', tmp_path/ 'root_path')
    monkeypatch.setattr('xls_management.WORKPATH', tmp_path / 'work_path')
    from xls_management.config import ATEConfig
    config = ATEConfig()
    assert isinstance(config.config_file, Path)
    assert isinstance(config.config, dict)

@parametrize_from_yaml(f'{MYDATA}/vw_dev/test_data/in/config_from.yml')
def test_config_config_from(given_config_path, expected, monkeypatch, tmp_path):
    del_xls_management_imports()
    monkeypatch.setattr('xls_management.ROOTPATH', tmp_path/ 'test/data')
    monkeypatch.setattr('xls_management.WORKPATH', tmp_path / 'work_path')
    from xls_management.config import ATEConfig
    config = ATEConfig()
    given_path = given_config_path.format(work_path=MYDATA)
    config.config_from(given_path)
    blacklist_attribute = config.get('blacklist_attribute', 'default')
    assert blacklist_attribute == expected

def test_config_config_from_invalid_path(monkeypatch, tmp_path):
    del_xls_management_imports()
    monkeypatch.setattr('xls_management.ROOTPATH', tmp_path/ 'test/data')
    monkeypatch.setattr('xls_management.WORKPATH', tmp_path / 'work_path')
    from xls_management.config import ATEConfig
    config = ATEConfig()
    with pytest.raises(FileNotFoundError):
        config.config_from("invalid\\path")
