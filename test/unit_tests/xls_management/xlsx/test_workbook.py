import pytest
from pathlib import Path
from shutil import copy

from test import working_path

import pandas as pd

from xls_management import WORKPATH
from xls_management.xlsx.workbook import Workbook


def test___init__(tmp_path) -> None:
    target_file = tmp_path / 'trial.xlsx'
    wb = Workbook(target_file)
    assert isinstance(wb, Workbook)
    assert wb.file_path == target_file

def test_writer(tmp_path) -> None:
    df = pd.DataFrame({"id": [1, 2], "values": ["Acuna\r\nMatata", "Lion\r\nking\r\nlives"]})
    target_file = tmp_path / 'trial.xlsx'
    assert target_file.exists() is False
    wb = Workbook(target_file)
    with wb.writer() as w:
        df.to_excel(
            w,
            sheet_name='Films',
            engine=wb.engine,
        )
    assert target_file.exists() is True

def test_reader() -> None:
    file_path: Path = working_path / "test/data/example.xlsx"
    wb = Workbook(file_path)
    kvargs={
        'sheet_name':'Cars',
        'dtype':str,
        'engine':wb.engine,
    }
    with wb.reader() as r:
        df = pd.read_excel(r, **kvargs)
    assert len(df) == 3
    assert 'License plate' in df.keys()
    assert df['License plate'][0] == 'DE2456HBZ'
    assert 'Brand' in df.keys()
    assert df['Brand'][2] == 'Audi'
    assert 'Modell' in df.keys()
    assert df['Modell'][1] == 'Polo' 

def test_append_worksheet(tmp_path) -> None:
    df = pd.DataFrame({"id": [1, 2], "values": ["Acuna\r\nMatata", "Lion\r\nking\r\nlives"]})
    target_file = tmp_path / 'trial.xlsx'
    assert target_file.exists() is False
    wb = Workbook(target_file)
    with wb.writer() as w:
        wb.append_worksheet(w, df, name='Films')
    assert target_file.exists() is True

    wb2 = Workbook(target_file)

    assert "Films" in wb2.sheet_names()
    
    df2:pd.DataFrame = wb2.sheet("Films")
    assert len(df2) == 2
    assert 'id' in df2.keys()
    assert df2['id'][0] == '1'
    assert 'values' in df2.keys()
    assert df2['values'][1] == 'Lion\nking\nlives'

def test_sheet_names():
    file_path: Path = working_path / "test/data/example.xlsx"
    w: Workbook = Workbook(file_path)
    result = w.sheet_names()
    assert 'People' in result
    assert 'Cars' in result
    assert len(result) == 2

def test_sheet():
    file_path: Path = working_path / "test/data/example.xlsx"
    wb: Workbook = Workbook(file_path)
    df:pd.DataFrame = wb.sheet('Cars')
    for name in ('License plate', 'Brand', 'Modell'):
        assert name in df.keys()
    
    assert len(df) == 3
    assert 'License plate' in df.keys()
    assert df['License plate'][0] == 'DE2456HBZ'
    assert 'Brand' in df.keys()
    assert df['Brand'][2] == 'Audi'
    assert 'Modell' in df.keys()
    assert df['Modell'][1] == 'Polo'

def test_all_sheets() -> None:
    file_path: Path = working_path / "test/data/example.xlsx"
    wb = Workbook(file_path)
    expected_length:dict[str, int] = {'People':3 ,'Cars':3}
    for sheet_name, df in wb.all_sheets():
        assert sheet_name in ('People', 'Cars')
        assert isinstance(df, pd.DataFrame)
        assert len(df) == expected_length[sheet_name]

def test_load_dataframe() -> None:
    file_path: Path = working_path / "test/data/example.xlsx"
    wb = Workbook(file_path)
    df:pd.DataFrame = wb.load_dataframe(sheet_name='Cars')
    assert len(df) == 3
    assert 'License plate' in df.keys()
    assert df['License plate'][0] == 'DE2456HBZ'
    assert 'Brand' in df.keys()
    assert df['Brand'][2] == 'Audi'
    assert 'Modell' in df.keys()
    assert df['Modell'][1] == 'Polo' 

def test_to_csv(tmp_path) -> None:
    file_path: Path = working_path / "test/data/example.xlsx"
    csv_file_path: Path = tmp_path / 'example_py.csv'
    workbook = Workbook(file_path)
    workbook.to_csv(
        csv_path= str(csv_file_path),
        sheet_name='Cars',
        slice_size = 0,
        mask_crlf=True,
    )
    assert csv_file_path.exists()

