from pathlib import Path
from test.conftest import parametrize_from_yaml

from pandas import DataFrame

from xls_management import WORKPATH
from xls_management.utils.compare_input import column_diff
from xls_management.xlsx.workbook import Workbook


@parametrize_from_yaml(file_path=f'{WORKPATH}\\vw\\test_data\\in\\compare_input.yml')
def test_compare_imput(
    key_name:str,
    columns:list[str],
    project_data:list[dict[str,str]],
    output_path:str,
):
    result = {'Fields':columns}
    for project_item in project_data:
        name, file, sheet_name = project_item.values()
        file_path = file.format(work_path=WORKPATH)
        assert Path(file_path).exists(), f'Unable to find {file_path}'
        w = Workbook(file_path)
        assert sheet_name in w.sheet_names(), f'{sheet_name} not in [{", ".join(w.sheet_names())}]'
        print(f'{name}')
        print(f'comparing {sheet_name} in {file_path}')
        diff = column_diff(w, sheet_name, key_name, output_path, columns)
        result[name] = [col['message'] for col in diff.values()]
    
    output_w = Workbook(output_path.format(work_path=WORKPATH))
    df = DataFrame(result)
    output_w.assure_parent_exists()
    with output_w.writer() as w:
        output_w.append_worksheet(w, df, name='Films')
    