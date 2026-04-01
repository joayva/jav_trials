from test import working_path

import pandas as pd
import pytest

from xls_management.is_ole import is_ole
from xls_management.xlsx.workbook import Workbook


def test_is_ole_non_existing():
    my_file = working_path / 'test/data/non_existing.xls'
    with pytest.raises(FileNotFoundError):
        assert is_ole(my_file) == f'{my_file} is ole'

def test_is_ole_empty_xlsx():
    my_file = working_path / 'test/data/Empty.xlsx'
    assert is_ole(my_file) == f'{my_file} is ole'

def test_pandas_xlsxwriter():
    example = working_path / 'test/data/example.xlsx'
    import pandas as pd

    # example data
    df1 = pd.DataFrame({
        'Name': ['Anna', 'Bernd', 'Clara'],
        'Age': [28, 34, 29],
        'City': ['Berlin', 'Hamburg', 'München']
    })
    
    df2 = pd.DataFrame({
        'License plate': ['DE2456HBZ', 'DE4562IBZ', 'DE5246ZHB'],
        'Brand': ['Volkswagen', 'Volkswagen', 'Audi'],
        'Modell': ['Passat', 'Polo', 'A4']
    })


    # Writing in an Excel worksheet
    with pd.ExcelWriter(example, engine='xlsxwriter') as writer:
        df1.to_excel(writer, sheet_name='People', index=False)

        # Zugriff auf das Workbook und Worksheet
        workbook = writer.book
        worksheet = writer.sheets['People']

        # Spaltenbreite setzen (A=0, B=1, C=2)
        worksheet.set_column(0, 0, 15)  # Name
        worksheet.set_column(1, 1, 8)   # Age
        worksheet.set_column(2, 2, 20)  # City

        df2.to_excel(writer, sheet_name='Cars', index=False)
        worksheet = writer.sheets['Cars']
        worksheet.set_column('A:A', 12)  # License plate
        worksheet.set_column('B:B', 20)  # Brand
        worksheet.set_column('C:C', 20)  # Modell

    assert is_ole(example) == f'{example} is ole'

def test_pandas_write_with_crlf():

    df2 = pd.DataFrame({'id': [1, 2], 'values': ['Test\r\nTest', 'pie\r\npie\r\npie']})
    with pd.ExcelWriter('prueba.xlsx') as w:
        df2.to_excel(w,index=False, engine='openpyxl', sheet_name='One',inf_rep='replace')
        # each \r value is stored as _x000D_
    with pd.ExcelFile('prueba.xlsx') as xls:
        df = pd.read_excel(xls, sheet_name='One')
        #fix: each value _x000D_ is replaced by \r
        df = df.replace(to_replace='_x000D_', value='\r', regex=True)
        assert df['values'][0] == 'Test\r\nTest'
        assert df['values'][1] == 'pie\r\npie\r\npie'