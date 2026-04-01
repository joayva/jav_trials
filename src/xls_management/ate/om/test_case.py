import re

import pandas as pd

from xls_management.ate.data_de import TestCaseAttribute
from xls_management.ate.om.verificationskriterium import Verificationskriterium
from xls_management.utils.tools import list_from_comma_separated_str


class TestCase:
#Option Explicit
#

    def __init__(
        self,
        #id:str,               #   Public TF_ID As String  'ID des Testfalls
        #name:str,             #   Public TF_Name As String    'Name des Testfalls
        #status:str,           #   Public TF_Status As String  'Status des Tesfalls
        #testinstanz:str,      #   Public TF_Testinstanz As String 'Testinstanz
        #testumgebungstyp:str, #   Public TF_Testumgebungstyp As String    'Testumgebungstyp
        #vk_id:str,            #   Public TF_VK_ID As String   'Testdesign-ID auf dem der Testfall basiert
        #anf_ids:list[str]=[], #   Public TF_anfIDs As Collection   'Sammlung der Anforderungen, 
                               #            die direkt oder indirekt mit dem Testfall verknüpft sind
        columns:pd.DataFrame,
        row:int,
        verification_criteria: dict[str,Verificationskriterium],
    ):
        ### moved from tracking
#       'TFs:   #1: ID,
                #2: Status, 
                #3: Testfallname, 
                #4: Sonstige-Varianten, 
                #5: Basierend auf Testdesign, 
                #6: verifiziert, 
                #7: Testinstanz
#       'Neuen Testfall anlegen
#       Set testfall = New Testfaelle
#       'Testfall-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       testfall.TF_ID = Replace(Replace(rngTFAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
        self.id = re.sub(r'[\?r]', '', str(columns['ID'][row]))
#       'Status Testfall einlesen
#       testfall.TF_Status = rngTFAttribute(2).Offset(lngZeile, 0).Value
        self.status = str(columns[TestCaseAttribute.Status][row])
#       'Testfall-Name einlesen
#       testfall.TF_Name = rngTFAttribute(3).Offset(lngZeile, 0).Value
        self.name = str(columns[TestCaseAttribute.TestCaseName][row])
#       'Testinstanz einlesen
#       testfall.TF_Testinstanz = rngTFAttribute(7).Offset(lngZeile, 0).Value
        self.test_instance =  str(columns[TestCaseAttribute.TestInstance][row])
#       'Testumgebungstyp einlesen
#       testfall.testumgebungstyp = Replace(rngTFAttribute(4).Offset(lngZeile, 0).Value, "Testumgebungstyp: ", "")
        self.test_environment_type =  re.sub(
            'Testumgebungstyp: ',
            '',
            str(columns[TestCaseAttribute.OtherVariants][row])
        )
#       
#       'Alle direkt mit dem aktuellen Testfall verknüpften Anforderungs-IDs erfassen
#       'direkte Testfälle nicht berücksichtigen!
#       'Anforderungs-IDs einlesen
#       'anfIDs = rngTFAttribute(6).Offset(lngZeile, 0).Value
        #requirement_ids = str(columns[TestCaseAttribute.Verified][row])
#       'Anforderungs-IDs nach Kommas trennen
#       'Set idList = EinlesenGetrennteWerteKomma(anfIDs)
#       'Anforderungs-IDs übernehmen
#       'Set testfall.TF_anfIDs = idList
        #self.requirement_ids = list_from_comma_separated_str(requirement_ids)
#       'Neue Sammlung für Anforderungs-IDs anlegen - Notwendig, wenn Liste der direkten Testfälle nicht übernommen wird
#       Set testfall.TF_anfIDs = New Collection
        self.requirement_ids = []
#       
#       'Alle über das Testdesign mit dem aktuellen Testfall verknüpften Anforderungs-IDs erfassen
#       'ID des übergeordneten Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       testfall.TF_VK_ID = Replace(Replace(rngTFAttribute(5).Offset(lngZeile, 0).Value, "?", ""), "r", "")
        self.vk_id = re.sub(r'[\?r]','',str(columns[TestCaseAttribute.TestDesignBased][row]))
#       'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
#       Set Verifikationskriterium = New Verifikationskriterium
#       Set Verifikationskriterium = FindeVK(verifikationKritList, testfall.TF_VK_ID)
        verification_criterion:Verificationskriterium|None = verification_criteria.get(self.vk_id,None)
#       If Not Verifikationskriterium Is Nothing Then
        if verification_criterion is not None:
#           'Anforderungs-ID aufnehmen
#           For Each varAnfID In Verifikationskriterium.anf_ids
            for requirement_id in verification_criterion.requirement_ids:
#               testfall.addElementID (varAnfID)
                self.add_element_id(requirement_id)
#           Next varAnfID
#           'Testfall aufnehmen
#           Verifikationskriterium.VK_Testfaelle.Add Item:=testfall, Key:=testfall.TF_ID
            verification_criterion.test_cases[self.id] = self
            verification_criteria[self.vk_id] = verification_criterion
#       End If

#   
#   Sub addElementID(ByVal elemID2 As String)
    def add_element_id(self, requirement_id:str):
#       Dim elemID1 As Variant
#       Dim isContained As Boolean
#       
#       isContained = False
#       For Each elemID1 In Me.TF_anfIDs
#           If (elemID1 = elemID2) Then
#               isContained = True
#               Exit For
#           End If
#       Next elemID1
#       If (isContained = False) Then
#           TF_anfIDs.Add elemID2
#       End If
        if not requirement_id in self.requirement_ids:
            self.requirement_ids.append(requirement_id)
#   End Sub
