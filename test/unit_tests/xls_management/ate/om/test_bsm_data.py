import re
from test import working_path
from test.conftest import parametrize_from_yaml

import pytest

from xls_management import WORKPATH
from xls_management.ate.data_de import RequirementAttribute
from xls_management.ate.om.bsm_data import BSMData
from xls_management.xlsx.workbook import Workbook


@parametrize_from_yaml(file_path = f'{WORKPATH}\\vw_dev\\test_data\\in\\bsm_data.yml')
def test___init__(input_data, expected) -> None:
    excel_path = input_data['excel_file'].format(work_path = WORKPATH)
    wb=Workbook(excel_path)
    req = wb.sheet('Rohdaten Conti')
    row = input_data['row']
    data = BSMData(
        columns=req,
        row=row,
        fru_timing_index=input_data['fruh_timing_index'],
        is_specific=input_data['is_specific'],
    )
    assert isinstance(data,BSMData)

    assert data.bsm_available == 'ja'
    assert data.fru_timing_index == input_data['fruh_timing_index']

    assert data.verifications_criteria == []
    assert data.test_cases == []
    assert data.avw_feature == expected[RequirementAttribute.Feature]
    assert data.avw_reifegrad == expected[RequirementAttribute.MaturityLevel]
    assert data.avw_implementer == expected[RequirementAttribute.Implementer]
    assert data.avw_dokument_id == expected[RequirementAttribute.DocumentID]
    assert data.avw_dokument_name == expected[RequirementAttribute.Document]
    assert data.avw_mv == expected[RequirementAttribute.MV]
    assert data.avw_id == expected[RequirementAttribute.ID]
    assert data.avw_status == expected[RequirementAttribute.Status]
    assert data.avw_typ == expected[RequirementAttribute.Type]
    assert data.avw_kategorie == expected[RequirementAttribute.Category]
    assert data.avw_bsm_safusi == expected[RequirementAttribute.BSMSaFuSiAssesment]
    assert data.avw_bsm_zz == expected[RequirementAttribute.BSMZZAssesment]
    assert data.avw_bsm_ed == expected[RequirementAttribute.BSMEDAssesment]
    assert data.avw_bsm_fff == expected[RequirementAttribute.BSMFFFAssesment]
    assert data.avw_bsm_o == expected[RequirementAttribute.BSMOAssesment]
    assert data.avw_bsm_se == expected[RequirementAttribute.BSMSeAssesment]
    assert data.avw_asil == expected[RequirementAttribute.ASIL]
    if (
        re.search('#abgelehnt_nicht_testbar', expected[RequirementAttribute.EditorialTeamComent], re.IGNORECASE) or 
        re.search('#abgelehnt_nicht_testbar', expected[RequirementAttribute.Temp1_Text], re.IGNORECASE)
    ):
        assert data.avw_abgelehnt_nicht_testbar == 'x'
    else:
        assert data.avw_abgelehnt_nicht_testbar == ''
    assert data.cluster_testing == expected[RequirementAttribute.TestingCluster]
    assert data.avw_anforderungsverantwortliche == expected[RequirementAttribute.RequirementOwners]

@pytest.mark.xfail   
def test_set_relevance() -> None:
    raise NotImplementedError("This test should be implemented")

@pytest.mark.xfail 
def test_set_i_stufe() -> None:
    raise NotImplementedError("This test should be implemented")

@pytest.mark.xfail 
def test_add_verification_criterion() -> None:
    raise NotImplementedError("This test should be implemented")

@pytest.mark.xfail 
def test_same_id() -> None:
    raise NotImplementedError("This test should be implemented")

@pytest.mark.xfail 
def test_requirement_data_to() -> None:
    raise NotImplementedError("This test should be implemented")
