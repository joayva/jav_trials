from pathlib import Path
from test.conftest import parametrize_from_yaml

import pandas as pd
import pytest

from xls_management import WORKPATH
from xls_management.utils.compare_output import (SHEET_NAME, compare_outputs,
                                                 load_sheet, normalize)
from xls_management.xlsx.workbook import Workbook


@parametrize_from_yaml(file_path=f'{WORKPATH}\\vw_dev\\test_data\\in\\compare_trials.yml')
def test_compare_td_status_outputs_trials(suffix:str, data_path:str, output_vba:str, output_py:str):
    data_root = data_path.format(workpath=WORKPATH)
    py_wb = Workbook(f"{data_root}\\{output_py}")
    vba_wb=Workbook(f"{data_root}\\{output_vba}")
    sheet_data = (
        (f'ATE_Status{suffix}','ID'),
        (f'TD_Status{suffix}','TD-VK'),
    )
    diff_wb = Workbook(f"{data_root}\\diff_Output.xlsx")
    diff_messages: list[str] = []
    with diff_wb.writer() as w:
        for sheet_name,key in sheet_data:
            py_df:pd.DataFrame = py_wb.load_dataframe(sheet_name=sheet_name)
            vba_df:pd.DataFrame = vba_wb.load_dataframe(sheet_name=sheet_name,skiprows=1)
            diff_df = compare_outputs(vba_df, py_df, key)
            if len(diff_df) > 0:
                diff_messages.append(f"{sheet_name}: {len(diff_df)} differences found")
            diff_wb.append_worksheet(w,diff_df,sheet_name)
    # differences should be 0
    assert len(diff_messages) == 0, '\n'.join(diff_messages)

def test_normalize_treats_empty_values_as_equivalent():
    series = pd.Series([None, float("nan"), "nan", "none", "<na>", "", "a\r\nb", "c\r", "_x000D_", "ok"])
    normalized = normalize(series)
    assert normalized.tolist() == ["", "", "", "", "", "", "a\nb", "c\n", "", "ok"]

def test_compare_outputs_returns_empty_for_equal_dataframes():
    df_vba = pd.DataFrame({"ID": [1, 2], "A": ["x", "y"], "B": [10, 20]})
    df_py = pd.DataFrame({"ID": [1, 2], "A": ["x", "y"], "B": [10, 20]})
    diff = compare_outputs(df_vba, df_py, key="ID")
    assert diff.empty

def test_compare_outputs_ignores_key_column_and_reports_differences():
    df_vba = pd.DataFrame({"ID": [1, 2], "A": ["x", "z"], "B": [10, 21]})
    df_py = pd.DataFrame({"ID": [1, 2], "A": ["x", "y"], "B": [10, 20]})
    diff = compare_outputs(df_vba, df_py, key="ID")

    assert len(diff) == 2
    assert diff["Column"].tolist() == ["A", "B"]
    assert diff["ID"].tolist() == [2, 2]
    assert diff["VBA"].tolist() == ["z", "21"]
    assert diff["Python"].tolist() == ["y", "20"]

def test_compare_outputs_normalizes_numeric_strings_and_boolean_strings():
    df_vba = pd.DataFrame({"ID": [1], "A": [1], "B": ["True"], "C": [None]})
    df_py = pd.DataFrame({"ID": [1], "A": ["1"], "B": ["True"], "C": [""]})
    diff = compare_outputs(df_vba, df_py, key="ID")
    assert diff.empty

def test_compare_outputs_uses_row_index_when_key_column_is_missing():
    df_vba = pd.DataFrame({"A": ["x"], "B": ["y"]})
    df_py = pd.DataFrame({"A": ["x"], "B": ["z"]})
    diff = compare_outputs(df_vba, df_py, key="ID")

    assert len(diff) == 1
    assert diff.iloc[0]["Row"] == 0
    assert diff.iloc[0]["ID"] == 0
    assert diff.iloc[0]["Column"] == "B"
    assert diff.iloc[0]["VBA"] == "y"
    assert diff.iloc[0]["Python"] == "z"

def test_load_sheet_reads_excel_file(tmp_path):
    excel_path = tmp_path / "compare_test.xlsx"
    expected = pd.DataFrame({"ID": [1, 2], "Name": ["alpha", "beta"]})
    expected.to_excel(excel_path, sheet_name=SHEET_NAME, index=False)

    loaded = load_sheet(excel_path, sheet_name=SHEET_NAME, header=0)
    pd.testing.assert_frame_equal(loaded, expected)

def test_load_sheet_raises_when_file_missing():
    with pytest.raises(FileNotFoundError):
        load_sheet(Path("nonexistent.xlsx"))
        