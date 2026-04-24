from xls_management.xlsx.workbook import Workbook


def get_duplicates_index(workbook: Workbook, sheet_name:str, key_name: str) -> dict[int|str,int]:
    """
    Find indices of duplicate rows based on a key column name.
    
    Args:
        workbook: Workbook object
        sheet_name: mame of the sheet to check
        key_name: Column name to check for duplicates
        
    Returns:
        List of row indices that contain duplicate values in the key column
    """
    index = {}
    df = workbook.sheet(sheet_name)
    i = 0
    for key_value in df[key_name]:
        if key_value in index.keys():
            index[key_value].append(i)
        else:
            index[key_value] = [i]
        i += 1
    
    return {key:value for key, value in index.items() if len(value)> 1}

def column_diff(workbook: Workbook, sheet_name:str, key_name: str, output_path:str, columns:list[str]=[]):
    dup_index = get_duplicates_index(workbook, sheet_name, key_name)
    df = workbook.sheet(sheet_name)
    if len(columns) == 0:
        columns = [col for col in df.columns if col != key_name]
    sheet_diff = {}
    for col in columns:
        col_diff = {}
        if col not in df.columns:
            msg =  f"❌ -----------"
        else:
            for key, rows in dup_index.items():
                value_0 = df[col][rows[0]]
                eq = [rows[0]]
                diff = []
                for row_i in rows[1:]:
                    if value_0 != df[col][row_i]:
                        diff.append((row_i,df[col][row_i]))
                    else:
                        eq.append(row_i)
                if len(diff) > 0:
                    col_diff[key] = {'rows': rows, 'reference':value_0, 'differences':diff,'matchs':eq}
            rows_count = len(dup_index)
            msg = ''
            if len(col_diff) == 0:
                msg = f'✅({rows_count}) {col}'
            else:
                msg = f'❌({len(col_diff)}/{rows_count}) {col}'
        sheet_diff[col] = {'differences':col_diff, 'message':msg}
        print(msg)
    return sheet_diff
