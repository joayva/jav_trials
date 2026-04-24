from pathlib import Path
from shutil import copy
from test import working_path

import pandas as pd

from xls_management import WORKPATH
from xls_management.xlsx.workbook import Workbook


def test_workbook_sheets():
    file_path: Path = working_path / "test/data/example.xlsx"
    w: Workbook = Workbook(file_path)
    result = w.sheet_names()
    assert 'People' in result
    assert 'Cars' in result
    assert len(result)

def test_workbook_sheet():
    file_path: Path = working_path / "test/data/example.xlsx"
    w: Workbook = Workbook(file_path)
    

def test_protected_workbook_sheets():
    file_path: Path = working_path / "test/data/example02.xlsx"
    w: Workbook = Workbook(file_path)
    result = w.sheet_names()
    assert result is None

def test_all_sheets():
    expected = {
        'People':('Name', 'Age', 'City'),
        'Cars':('License plate', 'Brand', 'Modell'),
    }

    for file_name in ("example.xlsx", "example01.xlsx"):
        file_path: Path = working_path / f"test/data/{file_name}"
        workbook:Workbook = Workbook(file_path)
        for name, df in workbook.all_sheets():
            assert name in ('People', 'Cars')
            for col in df.keys():
                assert col in expected[name]


def test_append_sheet(tmp_path):
    """Copy example01.xlsx to a temp location, append a sheet using pandas, and verify it appears."""
    src: Path = working_path / "test/data/example01.xlsx"
    dest: Path = tmp_path / "example01.xlsx"
    copy(src, dest)

    # create a simple DataFrame and append as a new sheet
    df = pd.DataFrame({"col1": [1, 2], "col2": ["a", "b"]})
    #with pd.ExcelWriter(dest, engine="openpyxl", mode="a") as writer:
    #    df.to_excel(writer, sheet_name="NewSheet", index=False)
    workbook = Workbook(file_path = dest)
    with workbook.writer() as writer:
        workbook.append_worksheet(writer, df, "NewSheet")

    # verify via Workbook wrapper
    w: Workbook = Workbook(dest)
    assert "NewSheet" in w.sheet_names()

    df2 = pd.DataFrame({"id": [1, 2], "values": ["Prueba\r\nPrueba", "pastel\r\npastel\r\npastel"]})

    # Cleanup is handled by tmp_path fixture, so no need to delete the file manually

def test_append_sheet_with_crlf(tmp_path):
    """Copy example01.xlsx to a temp location, append a sheet using pandas, and verify it appears."""
    src: Path = working_path / "test/data/example01.xlsx"
    dest: Path = tmp_path / "example01.xlsx"
    copy(src, dest)

    # create a simple DataFrame and append as a new sheet
    df = pd.DataFrame({"id": [1, 2], "values": ["Acuna\r\nMatata", "Lion\r\nking\r\nlives"]})
    #with pd.ExcelWriter(dest, engine="openpyxl", mode="a") as writer:
    #    df.to_excel(writer, sheet_name="NewSheet", index=False)
    w = Workbook(file_path = dest)
    with w.writer() as writer:
        w.append_worksheet(writer, df, "NewSheet")

    # verify via Workbook wrapper
    assert "NewSheet" in w.sheet_names()

    df = w.sheet("NewSheet")
    assert df['values'][0] == 'Acuna\nMatata'

    # Cleanup is handled by tmp_path fixture, so no need to delete the file manually

def test_excel_py_to_csv():
    workbook = Workbook(f"{WORKPATH}\\vw_dev\\data\\output.xlsx")
    workbook.to_csv(
        csv_path='{workdir}/csv/{sheet_name}/py/{sheet_name}.csv',
        sheet_name='TD_Status_013_2026',
        slice_size = 100
    )
    workbook.to_csv(
        csv_path='{workdir}/csv/{sheet_name}/py/{sheet_name}.csv',
        sheet_name='ATE_Status_013_2026',
        slice_size = 100
    )
    assert Path(f"{WORKPATH}\\vw_dev\\data\\csv\\TD_Status_013_2026\\py\\TD_Status_013_2026_000.csv").exists()
    assert Path(f"{WORKPATH}\\vw_dev\\data\\csv\\TD_Status_013_2026\\py\\TD_Status_013_2026_001.csv").exists()
    assert Path(f"{WORKPATH}\\vw_dev\\data\\csv\\TD_Status_013_2026\\py\\TD_Status_013_2026_041.csv").exists()

def test_excel_py_to_csv_mask_crlf():
    workbook = Workbook(f"{WORKPATH}\\vw_dev\\data\\output.xlsx")
    workbook.to_csv(
        csv_path='{workdir}/csv/{sheet_name}_py.csv',
        sheet_name='TD_Status_013_2026',
        slice_size = 0,
        mask_crlf=True,
    )
    workbook.to_csv(
        csv_path='{workdir}/csv/{sheet_name}_py.csv',
        sheet_name='ATE_Status_013_2026',
        slice_size = 0,
        mask_crlf=True,
    )
    assert Path(f"{WORKPATH}\\vw_dev\\data\\csv\\TD_Status_013_2026_py.csv").exists()
    assert Path(f"{WORKPATH}\\vw_dev\\data\\csv\\ATE_Status_013_2026_py.csv").exists()

def test_excel_vba_to_csv_mask_crlf():
    vba_workbook=Workbook(f"{WORKPATH}\\vw_dev\\data\\trial\\output_Status.xlsx")
    vba_workbook.to_csv(
        skiprows=1,
        csv_path=f'{WORKPATH}\\vw_dev\\data\\csv''\\{sheet_name}_vba.csv',
        sheet_name='TD_Status_013_2026',
        slice_size = 0,
        mask_crlf=True,
    )
    vba_workbook.to_csv(
        skiprows=1,
        csv_path=f'{WORKPATH}\\vw_dev\\data\\csv''\\{sheet_name}_vba.csv',
        sheet_name='ATE_Status_013_2026',
        slice_size = 0,
        mask_crlf=True,
    )

def test_excel_vba_to_csv():
    vba_workbook=Workbook(f"{WORKPATH}\\vw_dev\\data\\trial\\output_Status.xlsx")
    vba_workbook.to_csv(
        skiprows=1,
        csv_path=f'{WORKPATH}\\vw_dev\\data\\csv''\\{sheet_name}\\vba\\{sheet_name}.csv',
        sheet_name='TD_Status_013_2026',
        slice_size = 100
    )
    vba_workbook.to_csv(
        skiprows=1,
        csv_path=f'{WORKPATH}\\vw_dev\\data\\csv''\\{sheet_name}\\vba\\{sheet_name}.csv',
        sheet_name='ATE_Status_013_2026',
        slice_size = 100
    )
   