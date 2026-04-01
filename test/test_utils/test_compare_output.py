import re

import pandas as pd

from xls_management import HOMEPATH
from xls_management.utils.compare_output import compare_outputs
from xls_management.xlsx.workbook import Workbook


def test_compare_td_status_outputs():
    py_wb = Workbook(f"{HOMEPATH}\\vw\\data\\output.xlsx")
    vba_wb=Workbook(f"{HOMEPATH}\\vw\\data\\trial\\output_Status.xlsx")
    sheet_data = (
        ('ATE_Status_013_2026','ID'),
        ('TD_Status_013_2026','TD-VK'),
    )
    diff_wb = Workbook(f"{HOMEPATH}\\vw\\data\\diff_Output.xlsx")
    with diff_wb.writer() as w:
        for sheet_name,key in sheet_data:
            py_df:pd.DataFrame = py_wb.load_dataframe(sheet_name=sheet_name)
            vba_df:pd.DataFrame = vba_wb.load_dataframe(sheet_name=sheet_name,skiprows=1)
            diff_df = compare_outputs(vba_df, py_df, key)
            diff_wb.append_worksheet(w,diff_df,sheet_name)