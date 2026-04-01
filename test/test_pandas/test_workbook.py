from pathlib import Path
from test import working_path

import pandas as pd

from xls_management.xlsx.workbook import Workbook


def test_open():
    file_path: Path = working_path / "test/data/example.xlsx"
    df:pd.DataFrame = pd.read_excel(file_path)
    for name in ('Age', 'City', 'Name'):
        assert name in df.keys()

def test_sheet_names():
    file_path: Path = working_path / "test/data/example.xlsx"
    excel_file =pd.ExcelFile(file_path)
    for name in ('People', 'Cars'):
        assert name in excel_file.sheet_names

def test_sheet():
    file_path: Path = working_path / "test/data/example.xlsx"
    df:pd.DataFrame =pd.read_excel(file_path, sheet_name='Cars')
    for name in ('License plate', 'Brand', 'Modell'):
        assert name in df.keys()

def test_data_frame_concat(): 
    """
    Cars and Aidi worksheets are read from example_concat.xlsx
    Cars DataFrame should append all columns
    """
    file_path: Path = working_path / "test/data/example_concat.xlsx"
    cars_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='Cars')
    assert len(cars_df) == 2, 'cars worksheet should have 2 rows'
    audi_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='Audi')
    assert len(audi_df) == 2, 'audi worksheet should have 2 rows'
    assert len(cars_df.keys()) == 3, 'Cars worksheet should have 3 columns'
    assert len(audi_df.keys()) == 3, 'Audi worksheet should have 3 columns'
    for name in ('License plate', 'Brand', 'Modell'):
        assert name in cars_df.keys(), f"Cars columns should include {name}"
        assert name in audi_df.keys(), f"Audi columns should include {name}"
    cars_df = pd.concat([cars_df, audi_df], ignore_index=True)
    assert len(cars_df) == 4
    for name in ('License plate', 'Brand', 'Modell'):
        assert name in cars_df.keys(), f"Cars columns should include {name}"

def test_data_frame_headers(): 
    """
    Headers worksheet is read from example_concat.xlsx
    Headers DataFrame should contain 1 column and no rows
    """
    file_path: Path = working_path / "test/data/example_concat.xlsx"
    headers_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='Car_headers')
    assert len(headers_df) == 0, 'Car_headers worksheet should have 0 rows'
    assert len(headers_df.keys()) == 3, 'Car_headers worksheet should have 3 columns'
    for name in ('License plate', 'Brand', 'Modell'):
        assert name in headers_df.keys(), f"Car_headers columns should include {name}"

def test_data_frame_empty(): 
    """
    Empty worksheet is read from example_concat.xlsx
    Empty DataFrame should contain no columns and no rows
    """
    file_path: Path = working_path / 'test/data/example_concat.xlsx'
    empty_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='Empty')
    assert len(empty_df) == 0, 'Empty worksheet should have 0 rows'
    assert len(empty_df.keys()) == 0, 'Empty worksheet should have 0 columns'

def test_data_frame_row_offset(): 
    """
    row_offset worksheet is read from example_concat.xlsx
    offset DataFrame should contain 3 unamed columns and 7 rows
    """
    file_path: Path = working_path / 'test/data/example_concat.xlsx'
    offset_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='row_offset')
    assert len(offset_df) == 7, 'row_offset worksheet should have 6 rows'
    assert len(offset_df.keys()) == 3, 'row_offset worksheet should have 3 columns'
    for name in offset_df.keys():
        assert name[:8] == 'Unnamed:', 'expected "Unnamed: *" got "{name}"'
    for row in range(0,3):
        for name in offset_df.keys():
            assert pd.isna(offset_df[name][row]), f'Value at {name}[{row}] should be NAN'
    assert offset_df['Unnamed: 0'][3] == 'Name'
    assert offset_df['Unnamed: 1'][3] == 'Age'
    assert offset_df['Unnamed: 2'][3] == 'City'

def test_data_frame_column_offset(): 
    """
    column_offset worksheet is read from example_concat.xlsx
    offset DataFrame should contain 6 columns --3 unamed-- and 3 rows
    """
    file_path: Path = working_path / 'test/data/example_concat.xlsx'
    offset_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='column_offset')
    assert len(offset_df) == 3, 'column_offset worksheet should have 4 rows'
    assert len(offset_df.keys()) == 6, 'column_offset worksheet should have 6 columns'
    for i in range(0,3):
        name= f"Unnamed: {i}"
        assert name in offset_df.keys()
        for row, item in enumerate(offset_df[name]):
            assert pd.isna(item), f'each value in {name}[{row}] should be NAN'
    for name in ('Name', 'Age', 'City'):
        name in offset_df.keys()
    
    offset_df.iloc[:, 3:]
    for name in ('Name', 'Age', 'City'):
        name in offset_df.keys()

def test_data_frame_fix_row_offset(): 
    """
    row_offset worksheet is read from example_concat.xlsx
    offset DataFrame should contain 3 unamed columns and 7 rows
    """
    file_path: Path = working_path / 'test/data/example_concat.xlsx'
    offset_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='row_offset')
    assert len(offset_df) == 7, 'row_offset worksheet should have 7 rows'
    assert len(offset_df.keys()) == 3, 'row_offset worksheet should have 3 columns'
    for name in offset_df.keys():
        assert name[:8] == 'Unnamed:', 'expected "Unnamed: *" got "{name}"'
    for row in range(0,3):
        for name in offset_df.keys():
            assert pd.isna(offset_df[name][row]), f'Value at {name}[{row}] should be NAN'
    assert offset_df['Unnamed: 0'][3] == 'Name'
    assert offset_df['Unnamed: 1'][3] == 'Age'
    assert offset_df['Unnamed: 2'][3] == 'City'

    #fix
    start = offset_df.first_valid_index()
    assert start is not None
    offset_df.columns = offset_df.iloc[start]
    start += 1
    offset_df = pd.DataFrame(offset_df.iloc[start:].reset_index(drop=True))

    #check
    for name in ('Name', 'Age', 'City'):
        assert name in offset_df.keys()
    assert len(offset_df) == 3

def test_data_frame_fix_column_offset(): 
    """
    column_offset worksheet is read from example_concat.xlsx
    offset DataFrame should contain 6 columns --3 unamed-- and 3 rows
    """
    file_path: Path = working_path / 'test/data/example_concat.xlsx'
    offset_df:pd.DataFrame =pd.read_excel(file_path, sheet_name='column_offset')
    assert len(offset_df) == 3, 'column_offset worksheet should have 4 rows'
    assert len(offset_df.keys()) == 6, 'column_offset worksheet should have 6 columns'
    for i in range(0,3):
        name= f"Unnamed: {i}"
        assert name in offset_df.keys()
        for row, item in enumerate(offset_df[name]):
            assert pd.isna(item), f'each value in {name}[{row}] should be NAN'
    for name in ('Name', 'Age', 'City'):
        name in offset_df.keys()

    #fix
    offset_df = pd.DataFrame(offset_df.iloc[:, 3:])

    #check
    for name in ('Name', 'Age', 'City'):
        name in offset_df.keys()
    assert len(offset_df.keys()) == 3, 'DataFrame should have 3 columns now'
