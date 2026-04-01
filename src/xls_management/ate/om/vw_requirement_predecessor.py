import pandas as pd


#class AVWVorgaenger (DE)
class VWRequirementPredecessor: 
#Option Explicit
#

    def __init__(
        self,
        columns:pd.DataFrame,
        row:int,
    ):
        ###Moved from tracking.py
#       'Master-ID einlesen
#       AVWVorgaenger.ID = CStr(rngAVWMasterAttribute(1).Offset(lngZeile, 0).Value)
        self.id = str(columns['ID'][row])
#       'Kommentar Redaktionskreis und temp1_Text einlesen
#       If InStr(CStr(UCase(rngAVWMasterAttribute(2).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Or InStr(CStr(UCase(rngAVWMasterAttribute(3).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Then
        if(
            "#ABGELEHNT_NICHT_TESTBAR" in str(columns['temp1_Text'][row]).upper() or
            "#ABGELEHNT_NICHT_TESTBAR" in str(columns['Kommentar Redaktionskreis'][row]).upper()
        ):
#           AVWVorgaenger.AbgelehntNichtTestbar = "x"
            self.abgelehnt_nicht_testbar = "x"
#       End If