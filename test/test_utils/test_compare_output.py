import re
from test.conftest import parametrize_from_yaml

import pandas as pd

from xls_management import WORKPATH
from xls_management.utils.compare_output import compare_outputs
from xls_management.xlsx.workbook import Workbook


@parametrize_from_yaml(file_path=f'{WORKPATH}\\vw\\test_data\\in\\012_MEB21\\compare.yaml')
def test_compare_td_status_outputs(suffix:str, data_path:str, output_vba:str, output_py:str):
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
