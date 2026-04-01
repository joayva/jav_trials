 
# Define background colors
from openpyxl.styles import PatternFill
from xls_management.xlsx.fill_dict import FillDict
from xls_management.xlsx.colors import BG_GREEN,BG_YELLOW,BG_RED


def test_fill_dict():
    from xls_management.ate.data_de import OutputBSMAttribute

    fd:FillDict = FillDict()
    fd['John'][1] = BG_YELLOW
    assert len(fd['John']) == 1
    assert fd['John'][1] == BG_YELLOW
    fd['John'][2] = BG_RED
    assert len(fd['John']) == 2
    assert fd['John'][1] == BG_YELLOW
    assert fd['John'][2] == BG_RED
    fd[OutputBSMAttribute.ASIL.value][1]= BG_YELLOW
    assert 'ASIL' in fd.keys()
    colour:PatternFill = fd['ASIL'][1]
    assert colour == BG_YELLOW
    fd['John'][2] = BG_GREEN
    assert fd['John'][2] == BG_GREEN
