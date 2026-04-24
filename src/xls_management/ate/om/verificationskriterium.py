from __future__ import annotations

import re
from typing import TYPE_CHECKING

import pandas as pd

if TYPE_CHECKING:
    from xls_management.ate.om.test_case import TestCase

from xls_management.ate.data_de import \
    TDVCAttribute as VC  # Verification Criterion
from xls_management.ate.om.absicherungsauftraege import Absicherungsauftrag
from xls_management.utils.dict_first import DictFirst
from xls_management.utils.tools import list_from_comma_separated_str


class Verificationskriterium:
#Option Explicit
#

    def __init__(
        self,
    #    id:str,                #Public TF_ID As String  'ID des Testfalls
    #    name:str,              #Public TF_Name As String    'Name des Testfalls
    #    status:str,            #Public TF_Status As String  'Status des Tesfalls
    #    testinstanz:str,       #Public TF_Testinstanz As String 'Testinstanz
    #    testumgebungstyp:str,  #Public TF_Testumgebungstyp As String    'Testumgebungstyp
    #    vk_id:str,             #Public TF_VK_ID As String   'Testdesign-ID auf dem der Testfall basiert
    #    anf_ids:list[str]=[],  #Public TF_anfIDs As Collection   'Sammlung der Anforderungen, die direkt oder indirekt mit dem Testfall verknüpft sind
         columns:pd.DataFrame,
         row:int,
    ):
    #    self.id = id
    #    self.name = name
    #    self.status = status
    #    self.testinstanz = testinstanz
    #    self.testumgebungstyp = testumgebungstyp
    #    self.vk_id = vk_id
    #    self.anf_ids = anf_ids

###### From Sub EinlesenVerifikationskriterien() --initialization from a data_frame row
#       'ID des Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       strVerifikationsID = Replace(Replace(rngTDVKAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#       'ID des Verifikationsauftrags erfassen
#       verifikationKrit.VK_ID = strVerifikationsID
        self.vk_id = re.sub(r'[\?r]', '',str(columns[VC.ID][row]))
#       'Anforderungs-IDs einlesen
#       anfIDs = rngTDVKAttribute(2).Offset(lngZeile, 0).Value
        requirement_ids_str = str(columns[VC.RequirementBased][row])
#       'Anforderungs-IDs nach Kommas trennen
#       Set idList = EinlesenGetrennteWerteKomma(anfIDs)
#       'Alle mit dem aktuellen Verifikationskriterium verknüpften Anforderungs-IDs erfassen
#       Set verifikationKrit.anf_ids = idList
        self.requirement_ids = list_from_comma_separated_str(requirement_ids_str)
#       'Status des Verifikationskriteriums einlesen
#       verifikationKrit.VK_status = rngTDVKAttribute(3).Offset(lngZeile, 0).Value
        self.status = columns[VC.Status][row]
#       'Absicherungsaufträge für dieses Verifikationskriterium anlegen
#       Set verifikationKrit.Absicherungsauftraege = New Collection
        self.absicherungsauftraege:DictFirst[str,Absicherungsauftrag] = DictFirst()
#       'Sammlung für Testfälle vorbereiten
#       Set verifikationKrit.VK_Testfaelle = New Collection
        self.test_cases: DictFirst[str, "TestCase"] = DictFirst()
#       'Sammlung für I-Stufen vorbereiten
#       Set verifikationKrit.anf_IStufen = New Collection
        self.requirement_i_level:list[str] = []  
#       'Sammlung für Umsetzer vorbereiten
#       Set verifikationKrit.anf_Umsetzer = New Collection
        self.requirement_implementer:list[str] = []
#       'Sammlung für BsM-Relevanz vorbereiten
#       Set verifikationKrit.anf_BsMRelevanz = New Collection
        self.requirement_bsm_relevance:list[str] = []
#       'Sammlung für ASIL vorbereiten
#       Set verifikationKrit.anf_ASIL = New Collection
        self.requirement_asil:list[str] = []
#       'Sammlung für Feature vorbereiten
#       Set verifikationKrit.anf_Feature = New Collection
        self.requirement_feature:list[str] = []
#       'Sammlung für Reifegrad vorbereiten
#       Set verifikationKrit.anf_Reifegrad = New Collection
        self.requirement_maturity_level:list[str] = []
#       'Sammlung für Modulverantwortliche vorbereiten
#       Set verifikationKrit.anf_MV = New Collection
        self.requirement_mv:list[str] = []
#       'Sammlung für LAH-ID vorbereiten
#       Set verifikationKrit.anf_LAHID = New Collection
        self.requirement_lah_id:list[str] = []
#       'Sammlung für LAH-Namen vorbereiten
#       Set verifikationKrit.anf_LAHNamen = New Collection
        self.requirement_lah_name:list[str] = []
#       'Sammlung für Cluster Testing vorbereiten
#       Set verifikationKrit.anf_ClusterTesting = New Collection
        self.requirement_cluster_testing:list[str] = []
#       'Sammlung für Anforderungsverantwortliche vorbereiten
#       Set verifikationKrit.anf_Anforderungsverantwortliche = New Collection
        self.requirement_owner:list[str] = []
#       'Sammlung für Temp11_Auswahlfeld vorbereiten
#       Set verifikationKrit.anf_Temp11_Auswahlfeld = New Collection
        self.requirement_temp11_selection_field:list[str] = []
#       verifikationKrit.VK_temp1Text = rngTDVKAttribute(4).Offset(lngZeile, 0).Value
        self.temp1_text:str = columns[VC.Temp1Text][row]
#       'Aktion einlesen
#       verifikationKrit.VK_Aktion = rngTDVKAttribute(5).Offset(lngZeile, 0).Value
        self.aktion = str(columns[VC.Action][row])
        self.requirement_present = False

#   Sub addLAHName(ByVal elemName2 As String)
    def add_lah_name(self, name:str)-> None:
#       Dim elemName1 As Variant
#       Dim isContained As Boolean
#       
#       isContained = False
#       For Each elemName1 In Me.anf_LAHNamen
#           If (elemName1 = elemName2) Then
#               isContained = True
#               Exit For
#           End If
#       Next elemName1
#       If (isContained = False) Then
#           anf_LAHNamen.Add elemName2
#       End If
        if name not in self.requirement_lah_name:
            self.requirement_lah_name.append(name)
            
    def add_lah_id(self, id:str)-> None:
        if id not in self.requirement_lah_id:
            self.requirement_lah_id.append(id)
        
#   End Sub
#   
#   Sub addClusterTesting(ByVal elemName2 As String)
    def add_cluster_testing(self, name:str)->None:
#       Dim elemName1 As Variant
#       Dim isContained As Boolean
#       
#       If elemName2 = "" Then
#           elemName2 = "leer"
#       End If
#       
#       isContained = False
#       For Each elemName1 In Me.anf_ClusterTesting
#           If (elemName1 = elemName2) Then
#               isContained = True
#               Exit For
#           End If
#       Next elemName1
#       If (isContained = False) Then
#           anf_ClusterTesting.Add elemName2
#       End If
        if name not in self.requirement_cluster_testing:
            self.requirement_cluster_testing.append(name)
#   End Sub
    
