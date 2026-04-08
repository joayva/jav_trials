import pandas as pd

from xls_management.ate.data_de import OutputBSMAttribute
from xls_management.ate.om.bsm_data import BSMData
from xls_management.ate.om.fru_timming import FRUTiming
from xls_management.ate.om.verificationskriterium import Verificationskriterium
from xls_management.ate.om.vw_requirement_predecessor import \
    VWRequirementPredecessor


#class BSMNachfolgerDaten
class BSMSuccessorData(BSMData):
    def __init__(
        self,
        columns:pd.DataFrame,
        row:int,
        fru_timing_index: dict[str,FRUTiming],
        use_predecessor_ids:bool,
        predecessor_index:dict[str,VWRequirementPredecessor],
        is_specific: bool = False,
    ):
        super().__init__(columns, row, fru_timing_index, is_specific)
        if use_predecessor_ids:
            self.avw_predecessor_id = str(columns['ID der Vorgänger-Anforderung'][row])
#           'Kommentar Redaktionskreis und temp1_Text aus AVW-Vorgänger einlesen
#           Set AVWVorgaenger = New AVWVorgaenger
#           Set AVWVorgaenger = FindeAVWVorgaenger(AVWVorgaengerList, BSMDatensatz.AVWVorgaengerID)
            requirement_predecessor = self.predecessor_index.get(self.avw_predecessor_id, None)
#           If Not AVWVorgaenger Is Nothing Then
            if requirement_predecessor is not None:
#               If AVWVorgaenger.AbgelehntNichtTestbar = "x" Then
                if requirement_predecessor.abgelehnt_nicht_testbar == 'x':
#                   If BSMDatensatz.AVWAbgelehntNichtTestbar = "" Then
                    if self.avw_abgelehnt_nicht_testbar == '':
#                       BSMDatensatz.AVWAbgelehntNichtTestbar = "x (Master)"
                        self.avw_abgelehnt_nicht_testbar = 'x (Master)'
#                   Else
                    else:
#                       BSMDatensatz.AVWAbgelehntNichtTestbar = BSMDatensatz.AVWAbgelehntNichtTestbar & vbCrLf & "x (Master)"
                        self.avw_abgelehnt_nicht_testbar = f'{self.avw_abgelehnt_nicht_testbar}\n\rx (Master)'   
#                   End If
#               End If
#           End If
        else:
            self.avw_predecessor_id = None

    def same_id(self, requirement_id:id) -> bool:
        return requirement_id == self.avw_predecessor_id
    
    def requirement_data_to(self, row_output:dict[str,str]):
#       'Vorgänger ID
#       If blnAVWVorgaengerIDsVerwenden Then
#          rngBsMAttribute(0).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWVorgaengerID
        row_output[OutputBSMAttribute.RedirectedFrom] = self.avw_predecessor_id
        super().requirement_data_to(row_output)
        