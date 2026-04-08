import re

import pandas as pd

from xls_management.ate.data_de import (AVWProjectAttribute,
                                        OutputBSMAttribute,
                                        RequirementAttribute)
from xls_management.ate.om.fru_timming import FRUTiming
from xls_management.ate.om.test_case import TestCase
from xls_management.ate.om.verificationskriterium import Verificationskriterium


#class BSMDaten (DE)
class BSMData:
#Option Explicit
#

    def __init__(
        self,
        columns:pd.DataFrame,
        row:int,
        fru_timing_index: dict[str,FRUTiming],
        is_specific: bool = False,
    ):
        self.is_specific = is_specific
        self.bsm_available:str = 'ja'
        self.fru_timing_index = fru_timing_index
#       'Neues Verifikationskriterium anlegen
#       Set BSMDatensatz.Verifikationskriterium = New Collection
        self.verifications_criteria: list[Verificationskriterium] = []
#       'Neuen Testfall anlegen
#       Set BSMDatensatz.Testfaelle = New Collection
        self.test_cases: list[TestCase] = []
#       'Feature einlesen
#       BSMDatensatz.AVWFeature = CStr(rngAVWAttribute(7).Offset(lngZeile, 0).Value)
        self.avw_feature = str(columns[RequirementAttribute.Feature][row])
#       'Reifegrad einlesen
#       BSMDatensatz.AVWReifegrad = CStr(rngAVWAttribute(8).Offset(lngZeile, 0).Value)
        self.avw_reifegrad = str(columns[RequirementAttribute.MaturityLevel][row])
#       'Umsetzer einlesen
#       BSMDatensatz.AVWUmsetzer = CStr(rngAVWAttribute(9).Offset(lngZeile, 0).Value)
        self.avw_implementer = str(columns[RequirementAttribute.Implementer][row])
#       'Dokument-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       BSMDatensatz.AVWDokumentID = Replace(Replace(CStr(rngAVWAttribute(2).Offset(lngZeile, 0).Value), "?", ""), "r", "")
        self.avw_dokument_id = re.sub(r'[\?r]','',str(columns[RequirementAttribute.DocumentID][row]))
#       'Dokument-Name einlesen
#       BSMDatensatz.AVWDokumentName = CStr(rngAVWAttribute(19).Offset(lngZeile, 0).Value)
        self.avw_dokument_name = str(columns[RequirementAttribute.Document][row])
#       'Modulverantwortlichen einlesen
#       BSMDatensatz.AVWMV = CStr(rngAVWAttribute(17).Offset(lngZeile, 0).Value)
        self.avw_mv = str(columns[RequirementAttribute.MV][row])
#       'Anforderungs-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       BSMDatensatz.AVWID = Replace(Replace(CStr(rngAVWAttribute(1).Offset(lngZeile, 0).Value), "?", ""), "r", "")
        self.avw_id = re.sub(r'[\?r]','',str(columns[RequirementAttribute.ID][row]))
#       'Status einlesen
#       BSMDatensatz.AVWStatus = CStr(rngAVWAttribute(6).Offset(lngZeile, 0).Value)
        self.avw_status = str(columns[RequirementAttribute.Status][row])
#       'Typ einlesen
#       BSMDatensatz.AVWTyp = CStr(rngAVWAttribute(4).Offset(lngZeile, 0).Value)
        self.avw_typ = str(columns[RequirementAttribute.Type][row])
#       'Kategorie einlesen
#       BSMDatensatz.AVWKategorie = CStr(rngAVWAttribute(5).Offset(lngZeile, 0).Value)
        self.avw_kategorie = str(columns[RequirementAttribute.Category][row])
#       'BsM-Status einlesen
#       BSMDatensatz.AVWBsMSaFuSi = CStr(rngAVWAttribute(11).Offset(lngZeile, 0).Value)
        self.avw_bsm_safusi = str(columns[RequirementAttribute.BSMSaFuSiAssesment][row])
#       BSMDatensatz.AVWBsMZZ = CStr(rngAVWAttribute(12).Offset(lngZeile, 0).Value)
        self.avw_bsm_zz = str(columns[RequirementAttribute.BSMZZAssesment][row])
#       BSMDatensatz.AVWBsMED = CStr(rngAVWAttribute(13).Offset(lngZeile, 0).Value)
        self.avw_bsm_ed = str(columns[RequirementAttribute.BSMEDAssesment][row])
#       BSMDatensatz.AVWBsMFFF = CStr(rngAVWAttribute(14).Offset(lngZeile, 0).Value)
        self.avw_bsm_fff = str(columns[RequirementAttribute.BSMFFFAssesment][row])
#       BSMDatensatz.AVWBsMO = CStr(rngAVWAttribute(15).Offset(lngZeile, 0).Value)
        self.avw_bsm_o = str(columns[RequirementAttribute.BSMOAssesment][row])   
#       BSMDatensatz.AVWBsMSe = CStr(rngAVWAttribute(16).Offset(lngZeile, 0).Value)
        self.avw_bsm_se = str(columns[RequirementAttribute.BSMSeAssesment][row])
        self.set_relevance()
        ### refactored as BSMData.set_relevance()
#       'ASIL einlesen
#       BSMDatensatz.AVWASIL = CStr(rngAVWAttribute(10).Offset(lngZeile, 0).Value)
        self.avw_asil = str(columns[RequirementAttribute.ASIL][row])
#       'Kommentar Redaktionskreis und temp1_Text einlesen
#       If InStr(CStr(UCase(rngAVWAttribute(20).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Or InStr(CStr(UCase(rngAVWAttribute(21).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Then
        if (
            re.search('#abgelehnt_nicht_testbar', str(columns[RequirementAttribute.EditorialTeamComent][row]), re.IGNORECASE) or 
            re.search('#abgelehnt_nicht_testbar', str(columns[RequirementAttribute.Temp1_Text][row]), re.IGNORECASE)
        ):
#           BSMDatensatz.AVWAbgelehntNichtTestbar = "x"
            self.avw_abgelehnt_nicht_testbar = 'x'
        else:
            self.avw_abgelehnt_nicht_testbar = ''
#       End If
        self.set_i_stufe()
#       'Cluster Testing einlesen
#       BSMDatensatz.ClusterTesting = CStr(rngAVWAttribute(18).Offset(lngZeile, 0).Value)
        self.cluster_testing = str(columns[RequirementAttribute.TestingCluster][row])
#       
#       'Anforderungsverantwortliche einlesen
#       BSMDatensatz.AVWAnforderungsverantwortliche = CStr(rngAVWAttribute(22).Offset(lngZeile, 0).Value)
        self.avw_anforderungsverantwortliche = str(columns[RequirementAttribute.RequirementOwners][row])
#       
#       'Projekt MEB21 - Temp11_Auswahlfeld einlesen
#       If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
        if is_specific:
#           BSMDatensatz.AVWTemp11_Auswahlfeld = CStr(rngAVWAttributeMEB21(1).Offset(lngZeile, 0).Value)
            self.avw_temp11_auswahlfeld = str(columns[AVWProjectAttribute.Temp11SelectionField][row])
#       End If
    
    def set_relevance(self):
#       'Zusammenführung BsM-Relevanz
#       strBsMRelevanz = ""
        relevance = []
#       If CStr(BSMDatensatz.AVWBsMSaFuSi) = strBsMVorhanden Then
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-SaFuSi" Else strBsMRelevanz = strBsMRelevanz & ",BsM-SaFuSi"
        if self.avw_bsm_safusi == self.bsm_available:
#           strBsMRelevanz = "BsM-SaFuSi"
            relevance.append('BsM-SaFuSi')
#       End If
#       If CStr(BSMDatensatz.AVWBsMZZ) = strBsMVorhanden Then
        if self.avw_bsm_zz == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ZZ" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ZZ"
            relevance.append('BsM-ZZ')
#       End If
#       If CStr(BSMDatensatz.AVWBsMED) = strBsMVorhanden Then
        if self.avw_bsm_ed == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ED" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ED"
            relevance.append('BsM-ED')
#       End If
#       If CStr(BSMDatensatz.AVWBsMFFF) = strBsMVorhanden Then
        if self.avw_bsm_fff == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-FFF" Else strBsMRelevanz = strBsMRelevanz & ",BsM-FFF"
            relevance.append('BsM-FFF')
#       End If
#       If CStr(BSMDatensatz.AVWBsMO) = strBsMVorhanden Then
        if self.avw_bsm_o == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-O" Else strBsMRelevanz = strBsMRelevanz & ",BsM-O"
            relevance.append('BsM-O')
#       End If
#       If CStr(BSMDatensatz.AVWBsMSe) = strBsMVorhanden Then
        if self.avw_bsm_se == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-Se" Else strBsMRelevanz = strBsMRelevanz & ",BsM-Se"
            relevance.append('BsM-Se')
#       End If
#       BSMDatensatz.BSMRelevanz = strBsMRelevanz
        self.bsm_relevanz = ','.join(relevance)

    def set_i_stufe(self):
#       'Geplante I-Stufe einlesen
#       strIStufe = ""
        i_stufe = ''
#       strIStufeMin = ""
        min_i_stufe = ''
        if len(self.fru_timing_index) > 0:
#           If BSMDatensatz.AVWUmsetzer <> "" Then
            if self.avw_implementer != '':
#               varUmsetzer = Split(BSMDatensatz.AVWUmsetzer, ",", , vbBinaryCompare)
                converter_list = [converter.strip() for converter in self.avw_implementer.split(',')]     
#               For intUmsetzer = 0 To UBound(varUmsetzer, 1)
                for converter in converter_list:
#                   strIStufe = FRUTimingList.Item(BSMDatensatz.AVWFeature & BSMDatensatz.AVWReifegrad & Trim(varUmsetzer(intUmsetzer))).IStufe
                    fru_key = f'{self.avw_feature}{self.avw_reifegrad}{converter}'
                    try:
                        i_stufe = self.fru_timing_index[fru_key].i_stufe
                    except KeyError:
                        pass
#                   If strIStufeMin = "" Then
#                       strIStufeMin = strIStufe
#                   ElseIf InStr(strIStufe, "IS") > 0 Then
#                       If strIStufe < strIStufeMin Then
#                           strIStufeMin = strIStufe
#                       End If
#                   End If
                    if min_i_stufe == '' or ('IS' in i_stufe and i_stufe < min_i_stufe):
                        min_i_stufe = i_stufe
#               Next intUmsetzer
#           End If
#       BSMDatensatz.IStufe = strIStufeMin
        self.i_stufe = min_i_stufe
#

    def add_verification_criterion(self, verification_criterion:Verificationskriterium):
#       'VK-ID aufnehmen
#       BSMDatensatz.Verifikationskriterium.Add Item:=varErfassteVKItem 
        if verification_criterion not in self.verifications_criteria:
            self.verifications_criteria.append(verification_criterion)
#       'Geplante I-Stufe für Verifikationskritierum erfassen
#       If BSMDatensatz.IStufe <> "" Then
        if self.i_stufe != '' and self.i_stufe not in verification_criterion.requirement_i_level:
#           varErfassteVKItem.anf_IStufen.Add Item:=BSMDatensatz.IStufe
            verification_criterion.requirement_i_level.append(self.i_stufe)
#       End If
#       'Umsetzer für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWUmsetzer <> "" Then
        if self.avw_implementer != '' and self.avw_implementer not in verification_criterion.requirement_implementer:
#           varErfassteVKItem.anf_Umsetzer.Add Item:=BSMDatensatz.AVWUmsetzer
            verification_criterion.requirement_implementer.append(self.avw_implementer)
#       End If
#       'BsM-Relevanz für Verifikationskritierum erfassen
#       If BSMDatensatz.BSMRelevanz <> "" Then
        if self.bsm_relevanz != '' and self.bsm_relevanz not in verification_criterion.requirement_bsm_relevance:    
#           varErfassteVKItem.anf_BsMRelevanz.Add Item:=BSMDatensatz.BSMRelevanz
            verification_criterion.requirement_bsm_relevance.append(self.bsm_relevanz)
#       End If
#       'ASIL für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWASIL <> "" Then
        if self.avw_asil != '' and self.avw_asil not in verification_criterion.requirement_asil:
#           varErfassteVKItem.anf_ASIL.Add Item:=BSMDatensatz.AVWASIL
            verification_criterion.requirement_asil.append(self.avw_asil)
#       End If
#       'Feature für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWFeature <> "" Then
        if self.avw_feature != '' and self.avw_feature not in verification_criterion.requirement_feature:
#           varErfassteVKItem.anf_Feature.Add Item:=BSMDatensatz.AVWFeature
            verification_criterion.requirement_feature.append(self.avw_feature)
#       End If
#       'Reifegrad für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWReifegrad <> "" Then
        if self.avw_reifegrad != '' and self.avw_reifegrad not in verification_criterion.requirement_maturity_level:
#           varErfassteVKItem.anf_Reifegrad.Add Item:=BSMDatensatz.AVWReifegrad
            verification_criterion.requirement_maturity_level.append(self.avw_reifegrad)
#       End If
#       'Modulverantwortliche für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWMV <> "" Then
        if self.avw_mv != '' and self.avw_mv not in verification_criterion.requirement_mv:
#           varErfassteVKItem.anf_MV.Add Item:=BSMDatensatz.AVWMV
            verification_criterion.requirement_mv.append(self.avw_mv)
#       End If
#       'LAH-ID für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWDokumentID <> "" Then
        if self.avw_dokument_id != '' and self.avw_dokument_id not in verification_criterion.requirement_lah_id:
#           varErfassteVKItem.anf_LAHID.Add Item:=BSMDatensatz.AVWDokumentID
            verification_criterion.requirement_lah_id.append(self.avw_dokument_id)
#       End If
#       'LAH-Namen für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWDokumentName <> "" Then
        if self.avw_dokument_name != '' and self.avw_dokument_name not in verification_criterion.requirement_lah_name:
#           varErfassteVKItem.addLAHName (BSMDatensatz.AVWDokumentName)
            verification_criterion.add_lah_name(self.avw_dokument_name)
#       End If
#       'Cluster Testing für Verifikationskriterium erfassen
#       If BSMDatensatz.ClusterTesting <> "" Then
        if self.cluster_testing != '' and self.cluster_testing not in verification_criterion.requirement_cluster_testing:
#           varErfassteVKItem.anf_ClusterTesting.Add Item:=BSMDatensatz.ClusterTesting
            verification_criterion.requirement_cluster_testing.append(self.cluster_testing)
#       End If
#       'Anforderungsverantwortliche für Verifikationskriterium erfassen
#       If BSMDatensatz.AVWAnforderungsverantwortliche <> "" Then
        if (
            self.avw_anforderungsverantwortliche != '' and
            self.avw_anforderungsverantwortliche not in verification_criterion.requirement_owner
        ):   
#           varErfassteVKItem.anf_Anforderungsverantwortliche.Add Item:=BSMDatensatz.AVWAnforderungsverantwortliche
            verification_criterion.requirement_owner.append(self.avw_anforderungsverantwortliche)
#       End If
#       'Temp11_Auswahlfeld für Verifikationskriterium erfassen
#       If BSMDatensatz.AVWTemp11_Auswahlfeld <> "" Then
        if (
            self.is_specific and self.avw_temp11_auswahlfeld != '' and 
            self.avw_temp11_auswahlfeld not in verification_criterion.requirement_temp11_selection_field
        ):
#           varErfassteVKItem.anf_Temp11_Auswahlfeld.Add Item:=BSMDatensatz.AVWTemp11_Auswahlfeld
            verification_criterion.requirement_temp11_selection_field.append(self.avw_temp11_auswahlfeld)
#       End If
#       
#       'Innere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt

    def same_id(self, requirement_id:str) -> bool:
        return requirement_id == self.avw_id

    def requirement_data_to(self, row_output:dict[str,str]):
#       'Ausgabe ID
#       rngBsMAttribute(1).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWID
        row_output[OutputBSMAttribute.ID] = self.avw_id
#       'Ausgabe Dokument-ID
#       rngBsMAttribute(2).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWDokumentID
        row_output[OutputBSMAttribute.DocumentID] = self.avw_dokument_id
#       'Ausgabe BsM-Relevanz
#       rngBsMAttribute(3).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.BSMRelevanz
        row_output[OutputBSMAttribute.BSMRelevance] = self.bsm_relevanz
#       'Ausgabe BSM-SaFuSi
#       rngBsMAttribute(4).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMSaFuSi
        row_output[OutputBSMAttribute.BSMSaFuSiAssesment] = self.avw_bsm_safusi
#       'Ausgabe BSM-ZZ
#       rngBsMAttribute(5).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMZZ
        row_output[OutputBSMAttribute.BSMZZAssesment] = self.avw_bsm_zz
#       'Ausgabe BSM-ED
#       rngBsMAttribute(6).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMED
        row_output[OutputBSMAttribute.BSMEDAssesment] = self.avw_bsm_ed
#       'Ausgabe BSM-FFF
#       rngBsMAttribute(7).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMFFF
        row_output[OutputBSMAttribute.BSMFFFAssesment] = self.avw_bsm_fff
#       'Ausgabe BSM-O
#       rngBsMAttribute(8).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMO
        row_output[OutputBSMAttribute.BSMOAssesment] = self.avw_bsm_o
#       'Ausgabe BSM-Se
#       rngBsMAttribute(9).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMSe
        row_output[OutputBSMAttribute.BSMSeAssesment] = self.avw_bsm_se
#       'Ausgabe ASIL
#       rngBsMAttribute(10).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWASIL
        row_output[OutputBSMAttribute.ASIL] = self.avw_asil
#       'Ausgabe Feature
#       rngBsMAttribute(11).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWFeature
        row_output[OutputBSMAttribute.Feature] = self.avw_feature
#       'Ausgabe Reifegrad
#       rngBsMAttribute(12).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWReifegrad
        row_output[OutputBSMAttribute.MaturityLevel] = self.avw_reifegrad
#       'Ausgabe Umsetzer
#       rngBsMAttribute(13).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWUmsetzer
        row_output[OutputBSMAttribute.Implementer] = self.avw_implementer
#       'Ausgabe Status
#       rngBsMAttribute(14).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWStatus
        row_output[OutputBSMAttribute.Status] = self.avw_status
#       'Ausgabe MV
#       rngBsMAttribute(20).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWMV
        row_output[OutputBSMAttribute.MV] = self.avw_mv
#       'Ausgabe Kategorie
#       rngBsMAttribute(21).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWKategorie
        row_output[OutputBSMAttribute.Category] = self.avw_kategorie
#       'Ausgabe Dokumentenname
#       rngBsMAttribute(22).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWDokumentName
        row_output[OutputBSMAttribute.Document] = self.avw_dokument_name
#       'Ausgabe #abgelehnt_nicht_testbar
#       rngBsMAttribute(23).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWAbgelehntNichtTestbar
        row_output[OutputBSMAttribute.RejectedNotTestable] = self.avw_abgelehnt_nicht_testbar
        
#       'Ausgabe Zugeordnete I-Stufe
#       rngBsMAttribute(24).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.IStufe
        row_output[OutputBSMAttribute.AssignedILevel] = self.i_stufe
#       'Ausgabe Cluster Testing
#       rngBsMAttribute(29).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.ClusterTesting
        row_output[OutputBSMAttribute.TestingCluster] = self.cluster_testing
#       'Ausgabe Anforderungsverantwortliche
#       rngBsMAttribute(33).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWAnforderungsverantwortliche
        row_output[OutputBSMAttribute.RequirementOwner] = self.avw_anforderungsverantwortliche