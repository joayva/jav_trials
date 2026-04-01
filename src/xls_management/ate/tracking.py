from datetime import date, datetime
from pathlib import Path
import re
from itertools import islice

from openpyxl.styles import PatternFill
import pandas as pd

from xls_management.ate.om.bsm_data import BSMData
from xls_management.ate.om.bsm_successor_data import BSMSuccessorData
from xls_management.ate.om.db_info import DBInfo 
from xls_management.ate.om.fru_timming import FRUTiming
from xls_management.ate.om.project_db_info import ProjectDBInfo
from xls_management.ate.om.test_case import TestCase
from xls_management.ate.om.verificationskriterium  import Verificationskriterium
from xls_management.ate.om.vw_requirement_predecessor import VWRequirementPredecessor
from xls_management.ate.om.absicherungsauftraege import Absicherungsauftrag
from xls_management.ate.om.test_environment_evaluation import TestEnvironmentEvaluations

from xls_management.xlsx.fill_dict import FillDict
from xls_management.utils.tools import lazy_join as unic_join
#TODO: it should be replaced by unic_join after comparison with VBA script was fulfilled

from xls_management.ate.data_de import (
    RequirementAttribute,
    RequirementMasterAttribute,
    OutputBSMAttribute,
    TDProjectAttribute,
    TDSafeGuardsAttribute,
    TDVCAttribute,
    TestCaseAttribute,
        FRUTimingAttribute,
        TDAttribute,
    KNOWN_TEST_ENVIRONMENTS,
    RELEVANT_TEST_ENVIRONMENT_TOP as RELEVANT_TOP,
)
from xls_management.ate.project import project_combo_box
from xls_management.tui.msgbox import msgbox
from xls_management.config import ATEConfig
from xls_management.utils.tools import list_from_comma_separated_str

from importlib.metadata import version

from xls_management.xlsx.workbook import Workbook
from xls_management.xlsx.colors import BG_GREEN,BG_YELLOW,BG_RED

CRLF ='_x000D_\n'

#Option Explicit
#
#'Dokumentenabfrage einzeln \\vw.vwg\vwdfs\K-E\EF\1508\Groups\EFBS2_Konsulter\Testmanagement_EFDB\Projekt MQB48W\Testdesign\Statistik_TD_TS\
#
class ATEStatus:
#
#       Public Sub ATE_Status()
    def __init__(
        self,
    ) -> None:
        # set_signatures
        today = date.today()
        self.week_number = int(today.strftime('%W')) + 1
        self.date_signature = f"{today.year}/{self.week_number:02d}"
        self.date_suffix = f"{self.week_number:03d}_{today.year}"

        self.config=ATEConfig()
        self.use_predecessor_ids = False
        self.project = None
#       'Klasse Verifikationskriterien mit Absicherungsaufträgen
#       Public verifikationKritList As New Collection
        self.verification_criteria:dict = {}
#       'Klasse AVW-Rohdaten
#       Public BsMDatenList As New Collection
        self.bsm_datasets:dict = {}
#       'Klasse Testfälle
#       Public testfallList As New Collection
        self.test_cases:dict = {}
#       'Klasse FRU_Timing
#       Public FRUTimingList As New Collection
        self.fru_timming_index:dict = {}
#       'Klasse AVWVorgaenger
#       Public AVWVorgaengerList As New Collection
        self.predecessor_index_AVW:dict = {}
#       'Flag für die Berücksichtigung von Vorgänger-IDs bei den AVW-Rohdaten
#       Public blnAVWVorgaengerIDsVerwenden As Boolean
        self.use_predecessor_ids:bool
#       'BsM_Status
#       Dim wbBsM As Workbook                       'Workbook für BsM_Status
#       Dim wksBsM As Worksheet                     'Worksheet für BsM_Status
#       Dim strBsMAttribute() As String             'String-Array mit Attributen des Arbeitsblatts BsM_Status
#       Dim rngBsMAttribute() As Range              'Range-Array mit Attributen des Arbeitsblatts BsM_Status
        self.status_BsM:DBInfo
#       'TD_Status
#       Dim wbTD As Workbook                        'Workbook für TD_Status
#       Dim wksTD As Worksheet                      'Worksheet für TD_Status
#       Dim strTDAttribute() As String              'String-Array mit Attributen des Arbeitsblatts TD_Status
#       Dim rngTDAttribute() As Range               'Range-Array mit Attributen des Arbeitsblatts TD_Status
        self.status_TD:DBInfo
#       'AVW_Rohdaten - Projekt
#       Dim wbAVW As Workbook                       'Workbook für AVW_Rohdaten
#       Dim wksAVW As Worksheet                     'Worksheet für AVW_Rohdaten
#       Dim strAVWAttribute() As String             'String-Array mit Attributen des Arbeitsblatts AVW_Rohdaten
#       Dim rngAVWAttribute() As Range              'Range-Array mit Attributen des Arbeitsblatts AVW_Rohdaten
#       Dim strAVWAttributeMEB21() As String        'String-Array mit Attributen des Arbeitsblatts AVW_Rohdaten für MEB21
#       Dim rngAVWAttributeMEB21() As Range         'Range-Array mit Attributen des Arbeitsblatts AVW_Rohdaten für MEB21
        self.row_data_AVW:DBInfo|ProjectDBInfo
#       'AVW_Rohdaten - Master
#       Dim wbAVWMaster As Workbook                 'Workbook für AVWMaster_Rohdaten
#       Dim wksAVWMaster As Worksheet               'Worksheet für AVWMaster_Rohdaten
#       Dim strAVWMasterAttribute() As String       'String-Array mit Attributen des Arbeitsblatts AVWMaster_Rohdaten
#       Dim rngAVWMasterAttribute() As Range        'Range-Array mit Attributen des Arbeitsblatts AVWMaster_Rohdaten
        self.master_row_data_AVW:DBInfo
#       'TDVKs (Tesdesigns - Verifikationskriterium)
#       Dim wbTDVK As Workbook                      'Workbook für TDs - Verifikationskriterien
#       Dim wksTDVK As Worksheet                    'Worksheet für TDs - Verifikationskriterium
#       Dim strTDVKAttribute() As String            'String-Array mit Attributen des Arbeitsblatts TDs - Verifikationskriterium
#       Dim rngTDVKAttribute() As Range             'Range-Array mit Attributen des Arbeitsblatts TDs - Verifikationskriterium
        self.test_design_verification_criterion:DBInfo
#       'TDAAs (Tesdesigns - Absicherungsaufträge)
#       Dim wbTDAA As Workbook                      'Workbook für TDs - Absicherungsaufträge
#       Dim wksTDAA As Worksheet                    'Worksheet für TDs - Absicherungsaufträge
#       Dim strTDAAAttribute() As String            'String-Array mit Attributen des Arbeitsblatts TDs - Absicherungsaufträge
#       Dim rngTDAAAttribute() As Range             'Range-Array mit Attributen des Arbeitsblatts TDs - Absicherungsaufträge
        self.test_design_assurance_contracts:DBInfo
#       'TF (Testfälle)
#       Dim wbTF As Workbook                        'Workbook für Testfälle
#       Dim wksTF As Worksheet                      'Worksheet für Testfälle
#       Dim strTFAttribute() As String              'String-Array mit Attributen des Arbeitsblatts Testfälle
#       Dim rngTFAttribute() As Range               'Range-Array mit Attributen des Arbeitsblatts Testfälle
        self.test_case:DBInfo
#       'FRUTiming
#       Dim wbFRUTiming As Workbook                 'Workbook für FRU-Timing
#       Dim wksFRUTiming As Worksheet               'Worksheet für FRU-Timing
#       Dim strFRUTimingAttribute() As String       'String-Array mit Attributen des Arbeitsblattes für FRU-Timing
#       Dim rngFRUTimingAttribute() As Range        'Range-Array mit Attributen des Arbeitsblattes für FRU-Timing
        self.timming_FRU:DBInfo
        self.import_attribute = [False, False, False, False, False, False]
        self.errors:str = ''
#       'BsM-Status wird separat im Workbook des Makros erzeugt
#       Set wbBsM = ThisWorkbook
        file_path_BsM:str|None = self.config.get('workbook_path_BsM')
        assert file_path_BsM is not None
        self.workbook_BsM = Workbook(file_path=file_path_BsM)
        #self.output_workbook = Workbook(output_path)
    

    def perform_status(self, output_path:Path|None=None):
#       'Allgemein
#       Dim strFehlerGesamt As String               'String für Gesamtfehlerausgabe
#       Dim strWeitereTUsAusgabe As String          'String für Ausgabe weiterer Testumgebungstypen
#       Dim strLAHBlacklist() As String             'String-Array für einzulesende LAH-Blacklist
#       Dim strVersionMakro As String               'String-Array für Makro-Version
#       Dim strDateinamen(1 To 6) As String         'String-Array für die Namen der eingelesenen Dateien
#       Dim strProjekte(0 To 5) As String           'String-Array für die auswählbaren Fahrzeugprojekte
#       Dim strProjekt As String                    'String des ausgewählten Fahrzeugprojekts
#       'Verlauf
#       Dim strFehlerATEVerlauf As String           'String für Rückgabewert der Befüllung von ATE_Status_Verlauf
#       Dim strFehlerTDVerlauf As String            'String für Rückgabewert der Befüllung von TD_Status_Verlauf
#       Dim strFehlerVerlauf As String              'String für gemeinsamen Rückgabewert der Befüllung von ATE/TD_Status_Verlauf
#       
#       'Makro-Version
#       strVersionMakro = "ATE-Status V015F6" & vbCrLf & "Programmiert von Alexander Kuhlicke, Tagueri AG 2024"
#       
#       self.info_BsM = DBInfo()
#       'Befüllung der Projektliste
#       strProjekte(0) = "leer"
#       strProjekte(1) = "MQB48W"
#       strProjekte(2) = "MQB37W PA"
#       strProjekte(3) = "MEB UNECE"
#       strProjekte(4) = "MEB21"
#       strProjekte(5) = "Andere"
#       BoxAuswahlProjekt.ComboBox1.list() = strProjekte
#       BoxAuswahlProjekt.ComboBox1.ListIndex = 0
#       
#       'Abfrage Projekt und Nutzung Master-Bereich
#       BoxAuswahlProjekt.Caption = "ATE-Status " & strVersionMakro
#       BoxAuswahlProjekt.Show
        self.project, self.use_predecessor_ids = project_combo_box()
        self.is_project_specific = self.project in ['MEB21', 'MQB48W']
#       
#       If boolAuswahlGetroffen Then
#           'Eingabe Projekt
#           strProjekt = BoxAuswahlProjekt.ComboBox1.list(BoxAuswahlProjekt.ComboBox1.ListIndex)
#           'Eingabe Verwendung Master-IDs
#           If BoxAuswahlProjekt.OptionButton1.Value = True Then
#               blnAVWVorgaengerIDsVerwenden = True
#           Else
#               blnAVWVorgaengerIDsVerwenden = False
#           End If
#           
#           'Einlesen der Attribute der Rohdaten
#           If ATE_Status_Initializer(wbAVW, wbAVWMaster, wbTDVK, wbTDAA, wbTF, wbFRUTiming, _
#                                     wksAVW, wksAVWMaster, wksTDVK, wksTDAA, wksTF, wksFRUTiming, _
#                                     strAVWAttribute, strAVWMasterAttribute, strTDVKAttribute, strTDAAAttribute, strTFAttribute, strFRUTimingAttribute, _
#                                     rngAVWAttribute, rngAVWMasterAttribute, rngTDVKAttribute, rngTDAAAttribute, rngTFAttribute, rngFRUTimingAttribute, _
#                                     strFehlerGesamt, strDateinamen, blnAVWVorgaengerIDsVerwenden, _
#                                     strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21) Then
        if self.initialized():
#               'LAH-Blacklist einlesen
#               Call EinlesenLAHBlacklist(wbBsM, strLAHBlacklist)
            self.read_blacklist_LAHB()
#               'Testdesigns - Verifikationskriterien einlesen
#               Call EinlesenTDVKs(wksTDVK, strTDVKAttribute, rngTDVKAttribute)
            self.read_TDVKs()
#               'Testdesigns - Absicherungsaufträge einlesen
#               Call EinlesenTDAAs(wksTDAA, strTDAAAttribute, rngTDAAAttribute)
            self.read_TDAAs()
#               'Testfälle
#               Call EinlesenTFs(wksTF, strTFAttribute, rngTFAttribute)
            self.read_test_cases()
#               'FRU-Timing
#               Call EinlesenFRUTiming(wksFRUTiming, strFRUTimingAttribute, rngFRUTimingAttribute)
            self.read_FRU_timing()
#               'Anforderungen
                #coded in else side
#               If blnAVWVorgaengerIDsVerwenden = False Then
#                   'Anforderungsstatistik Projekt
#                   Call EinlesenAVWRohdaten(wksAVW, strAVWAttribute, rngAVWAttribute, strLAHBlacklist, strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21)
#               Else
            if self.use_predecessor_ids:
#                   'Anforderungsstatistik Masterbereich
#                   Call EinlesenAVWVorgaengerRohdaten(wksAVWMaster, strAVWMasterAttribute, rngAVWMasterAttribute)
                self.read_predecessor_requirement_raw_data()
#                   'Anforderungsstatistik Projekt
#                   Call EinlesenAVWNachfolgerRohdaten(wksAVW, strAVWAttribute, rngAVWAttribute, strLAHBlacklist, strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21)
                self.read_successor_requirement_raw_data()
#               End If
            else:
                self.read_requirement_raw_data()
#               'Ausgabe ATE-Status
#               Call AusgabeATEStatus(wbBsM, wksBsM, strBsMAttribute, rngBsMAttribute, strWeitereTUsAusgabe, strDateinamen, strProjekt)
#               'Ausgabe TD-Status
#               Call AusgabeTDStatus(wbBsM, wksTD, strTDAttribute, rngTDAttribute, strDateinamen, strProjekt)
#               'Geöffnete Dateien schliessen
#               Call SchliessenWb(wbBsM, wbAVW, wbAVWMaster, wbTDVK, wbTDAA, wbTF, wbFRUTiming)
                if output_path is None:
                    output_path = self.workbook_BsM.file_path.parent / "output.xlsx"
                wb = Workbook(output_path)
                with wb.writer() as writer:
                    for row_data_set, name, bg_colours in self.output_worksheets():
                        try:
                            df:pd.DataFrame = pd.DataFrame(
                                row_data_set,
                                dtype=str
                            )
                            wb.append_worksheet(writer, df, name, bg_colours=bg_colours)
                        except PermissionError:
                            print("Error: The file is open in another program. Please close it and try again.")
                        except Exception as  e:
                            print(f"An error occurred: {e}")    
            
#               
#               'Verläufe ATE_Status_Verlauf und TD_Status_Verlauf befüllen und Rückgabewerte zusammenführen
#               Call AusgabeVerlauf(wksBsM, strFehlerATEVerlauf, 1)
#               Call AusgabeVerlauf(wksTD, strFehlerTDVerlauf, 2)
            self.output_history()
            #TODO: errors seems to be join together at next, so taking them together in a self.errors instance
            #     at ATEStatus object seems to be the right thing
            #
#               If strFehlerATEVerlauf <> "" And strFehlerTDVerlauf <> "" Then
#                   strFehlerVerlauf = strFehlerATEVerlauf & vbCrLf & strFehlerTDVerlauf
#               ElseIf strFehlerATEVerlauf <> "" Then
#                   strFehlerVerlauf = strFehlerATEVerlauf
#               ElseIf strFehlerTDVerlauf <> "" Then
#                   strFehlerVerlauf = strFehlerTDVerlauf
#               Else
#                   strFehlerVerlauf = ""
#               End If
            ###################################
#               
#               'Abschlussmeldung
#               If strWeitereTUsAusgabe = "" Then
            if self.other_test_environment_output == "":
#                   If strFehlerVerlauf = "" Then
                if self.errors == "":
#                       MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
                    msgbox(
                        f"ATE-Status was created!\n\n"
                        f"{version('xls_management')}\n"
                    )
#                   Else
                else:
                    msgbox(
                        f"ATE-Status was created!\n\n"
                        f"{self.errors}\n\n"
                        "-----\n"
                        f"{version('xls_management')}\n"
                    )
#                       MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & strFehlerVerlauf & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
#                   End If
#               Else
#                   If strFehlerVerlauf = "" Then
            elif self.errors == "":
#                       MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & "Folgende weitere Testumgebungstypen wurden erkannt, aber nicht für den Vergleich berücksichtigt:" & vbCrLf & vbCrLf & strWeitereTUsAusgabe & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
                msgbox(
                    f"ATE-Status was created!\n\n"
                    "Folgende weitere Testumgebungstypen wurden erkannt, aber nicht für den Vergleich berücksichtigt:"
                    f"{self.other_test_environment_output}\n\n"
                    "-----\n"
                    f"{version('xls_management')}\n"
                ) 
#                   Else
            else:
#                       MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & "Folgende weitere Testumgebungstypen wurden erkannt, aber nicht für den Vergleich berücksichtigt:" & vbCrLf & vbCrLf & strWeitereTUsAusgabe & vbCrLf & vbCrLf & strFehlerVerlauf & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
                msgbox(
                    f"ATE-Status was created!\n\n"
                    "Folgende weitere Testumgebungstypen wurden erkannt, aber nicht für den Vergleich berücksichtigt:"
                    f"{self.other_test_environment_output}\n\n"
                    f"{self.errors}\n\n"
                    "-----\n"
                    f"{version('xls_management')}\n"
                )
#                   End If
#               End If
#               
#           Else
        else:
#               MsgBox strFehlerGesamt, Buttons:=vbExclamation, Title:="Fehler beim Import für ATE-Tracking"
            msgbox(
                f"Errors trying to import for ATE-Tracking\n\n"
                f"{self.errors}\n\n"
            )
#           End If
#           
#           'Worksheets und Klassenmodule zurücksetzen
#           Call ATE_Status_Deinitializer(wksBsM, wksAVW, wksTDVK, wksTDAA, wksTF, wksFRUTiming)
        self.status_deinitialize()
#       End If
#       
#       Unload BoxAuswahlProjekt
#       End Sub
#
    def initialized(self) -> bool:
#   Private Function ATE_Status_Initializer(ByRef wbAVW As Workbook, ByRef wbAVWMaster As Workbook, ByRef wbTDVK As Workbook, ByRef wbTDAA As Workbook, ByRef wbTF As Workbook, ByRef wbFRUTiming As Workbook, _
#           ByRef wksAVW As Worksheet, ByRef wksAVWMaster As Worksheet, ByRef wksTDVK As Worksheet, ByRef wksTDAA As Worksheet, ByRef wksTF As Worksheet, ByRef wksFRUTiming As Worksheet, _
#           ByRef strAVWAttribute() As String, ByRef strAVWMasterAttribute() As String, ByRef strTDVKAttribute() As String, ByRef strTDAAAttribute() As String, ByRef strTFAttribute() As String, ByRef strFRUTimingAttribute() As String, _
#           ByRef rngAVWAttribute() As Range, ByRef rngAVWMasterAttribute() As Range, ByRef rngTDVKAttribute() As Range, ByRef rngTDAAAttribute() As Range, ByRef rngTFAttribute() As Range, ByRef rngFRUTimingAttribute() As Range, _
#           ByRef strFehlerGesamt As String, ByRef strDateinamen() As String, ByVal blnAVWVorgaengerIDsVerwenden As Boolean, _
#           ByVal strProjekt As String, ByRef strAVWAttributeMEB21() As String, ByRef rngAVWAttributeMEB21() As Range) As Boolean
#           
#       Dim strFehlerBsM As String                  'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerAVW As String                  'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerAVWMaster As String            'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerTDAA As String                 'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerTDVK As String                 'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerTF As String                   'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerFRUTiming As String            'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strAttributeAVW As String               'String für alle Attribute
#       Dim strAttributeAVWMaster As String         'String für alle Attribute
#       Dim strAttributeTDAA As String              'String für alle Attribute
#       Dim strAttributeTDVK As String              'String für alle Attribute
#       Dim strAttributeTF As String                'String für alle Attribute
#       Dim strAttributeFRUTiming As String         'String für alle Attribute
#       Dim blnImportAttribute(1 To 6) As Boolean   'Flag-Array für korrektes Einlesen der Arbeitsblätter #1: AVW-Rohdaten, #2: TDVKs, #3: TDAAs, #4: Testfälle, #5: FRU-Timing, #6: AVWMaster-Rohdaten
#       Dim i As Integer                            'Laufvariable
#       
#       strFehlerGesamt = ""
        self.strFehlerGesamt = ""
#       
#       'Arbeitsblatt AVW_Rohdaten
#       'AVW: #1: ID, #2: Dokument-ID, #3: Basis für Testdesign, #4: Typ, #5: Kategorie, #6: Status, #7: Feature, #8: Reifegrad, #9: Umsetzer, #10: ASIL
#       '#11: BSM-SaFuSi Bewertung, #12: BSM-ZZ Bewertung, #13: BSM-ED Bewertung, #14: BSM-FFF Bewertung, #15: BSM-O Bewertung, #16: BSM-Se Bewertung, #17: MV AVWMV
#       '#18: Cluster Testing, #19: Dokument, #20: Kommentar Redaktionskreis, #21: temp1_Text, #22: Abgezweigt aus (vorher: Vorgänger ID, nur bei Projekt+Master)
#       'Weiche für Erfassung der Nachfolger-IDs
#       If blnAVWVorgaengerIDsVerwenden = False Then
#           ReDim strAVWAttribute(1 To 22)
#       Else
#           ReDim strAVWAttribute(1 To 23)
#       End If
#       ReDim rngAVWAttribute(LBound(strAVWAttribute, 1) To UBound(strAVWAttribute, 1))
#       If blnAVWVorgaengerIDsVerwenden Then
#           strAVWAttribute(23) = "Abgezweigt aus"  'strAVWAttribute(22) = "Abgezweigt aus"
#       End If
        if self.use_predecessor_ids:
            self.info_AVW = DBInfo(attributes=RequirementAttribute)
        else:
            self.info_AVW = DBInfo(
                attributes=islice(RequirementAttribute,0,len(RequirementAttribute)-1)
            )
#
#       'Dateiauswahl und Zuordnung
#       'Projektspezifisch (MEB21 oder MQB48W) oder allgemein
#       If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
        if self.project in ("MEB21", "MQB48W"):
            # Project specific importation
#           ReDim strAVWAttributeMEB21(1 To 1)
#           ReDim rngAVWAttributeMEB21(1 To 1)
#           strAVWAttributeMEB21(1) = "Temp11_Auswahlfeld"
            attributes_MBE21_AVW:tuple[str] = ('Temp11_Auswahlfeld',)
            self.info_AVW = ProjectDBInfo(
                path=self.config.get('default_path', '.'),
                db_info=self.info_AVW,
                project=self.project,
                project_attributes=attributes_MBE21_AVW,
            )
#           If EinlesenDatei_Projektspezifisch("Anforderungen Projekt " & strProjekt, strAVWAttribute, rngAVWAttribute, wbAVW, wksAVW, strFehlerAVW, strDateinamen(1), strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21) Then
#               blnImportAttribute(1) = True
            self.import_attribute[0] = self.info_AVW.einlesen_datei(f"Anforderungen Projekt {self.project}")
#           Else
#               'Sammlung aller gesuchten allgemeinen Attribute erzeugen
#               For i = LBound(strAVWAttribute, 1) To UBound(strAVWAttribute, 1)
#                   If strAttributeAVW = "" Then
#                       strAttributeAVW = strAVWAttribute(i)
#                   Else
#                       strAttributeAVW = strAttributeAVW & ", " & strAVWAttribute(i)
#                   End If
#               Next
            if not self.import_attribute[0]:
                #strAttributeAVW = self.info_AVW.str_attributes()
#               'Sammlung aller gesuchten projektspezifischen Attribute erzeugen
#               For i = LBound(strAVWAttributeMEB21, 1) To UBound(strAVWAttributeMEB21, 1)
#                   If strAttributeAVW = "" Then
#                       strAttributeAVW = strAVWAttributeMEB21(i)
#                   Else
#                       strAttributeAVW = strAttributeAVW & ", " & strAVWAttributeMEB21(i)
#                   End If
#               Next i
#               'Zusammenführung der gesuchten Attribute
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strProjekt & " - " & strAttributeAVW & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strProjekt & " - " & strAttributeAVW & ")"
#               End If
#               blnImportAttribute(1) = False
                self.errors += self.info_AVW.get_errors('Anforderungen können nicht eingelesen werden!')
#           End If
#       ElseIf EinlesenDatei("Anforderungen Projektbereich", strAVWAttribute, rngAVWAttribute, wbAVW, wksAVW, strFehlerAVW, strDateinamen(1)) Then
#           ImportAttribute(1) = True
        else:
            #No project specific importation
            self.import_attribute[0] = self.info_AVW.einlesen_datei("Anforderungen Projektbereich")
        if not self.import_attribute[0]:
#           'Sammlung aller gesuchten Attribute erzeugen
#           strAttributeAVW = ""
#           For i = LBound(strAVWAttribute, 1) To UBound(strAVWAttribute, 1)
#               If strAttributeAVW = "" Then
#                   strAttributeAVW = strAVWAttribute(i)
#               Else
#                   strAttributeAVW = strAttributeAVW & ", " & strAVWAttribute(i)
#               End If
#           Next i
            strAttributeAVW = self.info_AVW.str_attributes()
#           If strFehlerGesamt = "" Then
#               strFehlerGesamt = "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVW & ")"
#           Else
#               strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVW & ")"
#           End If
            self.errors += self.info_AVW.get_errors('Anforderungen können nicht eingelesen werden!')
#       End If
#       
#       If blnImportAttribute(1) Then
#           'Arbeitsblatt TDs - Verifikationskriterium
#           'TDVKs: #1: ID, #2: Basierend auf der Anforderung, #3: Status, #4: Temp1_Text, #5: Aktion
#           ReDim strTDVKAttribute(1 To 5)
#           ReDim rngTDVKAttribute(LBound(strTDVKAttribute, 1) To UBound(strTDVKAttribute, 1))
#           strTDVKAttribute(1) = "ID"
#           strTDVKAttribute(2) = "Basierend auf der Anforderung"
#           strTDVKAttribute(3) = "Status"
#           strTDVKAttribute(4) = "Temp1_Text"
#           strTDVKAttribute(5) = "Aktion"
        if self.import_attribute[0]:
            self.info_TDVK = DBInfo(
                path=self.config.get('default_path', '.'),
                attributes = TDVCAttribute,
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("Verifikationskriterien", strTDVKAttribute, rngTDVKAttribute, wbTDVK, wksTDVK, strFehlerTDVK, strDateinamen(2)) Then
#               blnImportAttribute(2) = True
            self.import_attribute[1] = self.info_TDVK.einlesen_datei("Verifikationskriterien")
#           Else
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeTDVK = ""
#               For i = LBound(strTDVKAttribute, 1) To UBound(strTDVKAttribute, 1)
#                   If strAttributeTDVK = "" Then
#                       strAttributeTDVK = strTDVKAttribute(i)
#                   Else
#                       strAttributeTDVK = strAttributeTDVK & ", " & strTDVKAttribute(i)
#                   End If
#               Next i
            if not self.import_attribute[1]:
                # str_attribute_TDVK = self.info_TDVK.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Verifikationskriterien können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDVK & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Verifikationskriterien können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDVK & ")"
#               End If
                self.errors += self.info_TDVK.get_errors('Verifikationskriterien können nicht eingelesen werden!')
#               blnImportAttribute(2) = False
#           End If
#       End If
#       
#       If blnImportAttribute(2) Then
#           'Arbeitsblatt TDs - Absicherungsaufträge
#           'TDAAs: #1: ID, #2: Enthalten in, #3: Status, #4: Testinstanz, #5: Testumgebungstyp
#           ReDim strTDAAAttribute(1 To 5)
#           ReDim rngTDAAAttribute(LBound(strTDAAAttribute, 1) To UBound(strTDAAAttribute, 1))
        if self.import_attribute[1]:
#           strTDAAAttribute(1) = "ID"
#           strTDAAAttribute(2) = "Enthalten in"
#           strTDAAAttribute(3) = "Status"
#           strTDAAAttribute(4) = "Testinstanz"
#           strTDAAAttribute(5) = "Testumgebungstyp"
            self.info_TDAA = DBInfo(
                path=self.config.get('default_path', '.'),
                attributes = TDSafeGuardsAttribute
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("Absicherungsaufträge", strTDAAAttribute, rngTDAAAttribute, wbTDAA, wksTDAA, strFehlerTDAA, strDateinamen(3)) Then
#               blnImportAttribute(3) = True
            self.import_attribute[2] = self.info_TDAA.einlesen_datei("Absicherungsaufträge")
#           Else
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeTDAA = ""
#               For i = LBound(strTDAAAttribute, 1) To UBound(strTDAAAttribute, 1)
#                   If strAttributeTDAA = "" Then
#                       strAttributeTDAA = strTDAAAttribute(i)
#                   Else
#                       strAttributeTDAA = strAttributeTDAA & ", " & strTDAAAttribute(i)
#                   End If
#               Next i
            if not self.import_attribute[2]:
                #str_attribute_TDAA = self.info_TDAA.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Absicherungsaufträge können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDAA & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Absicherungsaufträge können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDAA & ")"
#               End If
                self.errors += self.info_TDAA.get_errors('Absicherungsaufträge können nicht eingelesen werden!')
#               blnImportAttribute(3) = False
#           End If
#       End If
#         
#       If blnImportAttribute(3) Then
#           'Arbeitsblatt Testfälle
#           'TFs: #1: ID, #2: Status, #3: Testfallname, #4: Sonstige-Varianten, #5: Basierend auf Testdesign, #6: verifiziert, #7: Testinstanz
#           ReDim strTFAttribute(1 To 7)
#           ReDim rngTFAttribute(LBound(strTFAttribute, 1) To UBound(strTFAttribute, 1))
#           strTFAttribute(1) = "ID"
#           strTFAttribute(2) = "Status"
#           strTFAttribute(3) = "Testfallname"
#           strTFAttribute(4) = "Sonstige-Varianten"
#           strTFAttribute(5) = "Basierend auf Testdesign"
#           strTFAttribute(6) = "verifiziert"
#           strTFAttribute(7) = "Testinstanz"
        if self.import_attribute[2]:
            self.info_TF = DBInfo(
                path=self.config.get('default_path', '.'),
                attributes = TestCaseAttribute,
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("Testfälle", strTFAttribute, rngTFAttribute, wbTF, wksTF, strFehlerTF, strDateinamen(4)) Then
#               blnImportAttribute(4) = True
            self.import_attribute[3] = self.info_TF.einlesen_datei("Testfälle")
#           Else
            if not self.import_attribute[3]:
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeTF = ""
#               For i = LBound(strTFAttribute, 1) To UBound(strTFAttribute, 1)
#                   If strAttributeTF = "" Then
#                       strAttributeTF = strTFAttribute(i)
#                   Else
#                       strAttributeTF = strAttributeTF & ", " & strTFAttribute(i)
#                   End If
#               Next i
                #str_attribute_TF = self.info_TF.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Testfälle können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTF & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Testfälle können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTF & ")"
#               End If
                self.errors += self.info_TF.get_errors("Testfälle können nicht eingelesen werden!")
#               blnImportAttribute(4) = False
#           End If
#       End If
#                   
#       If blnImportAttribute(4) Then
        if self.import_attribute[3]:
#           'Arbeitsblatt FRU-Timing
#           'FRUTiming: #1: FeatureName, #2: RG, #3: Umsetzer, #4: Zuordnung zu I-Stufe
#           ReDim strFRUTimingAttribute(1 To 4)
#           ReDim rngFRUTimingAttribute(LBound(strFRUTimingAttribute, 1) To UBound(strFRUTimingAttribute, 1))
#           strFRUTimingAttribute(1) = "FeatureName"
#           strFRUTimingAttribute(2) = "Reifegrad"  'vorher "RG"
#           strFRUTimingAttribute(3) = "Umsetzer"
#           strFRUTimingAttribute(4) = "FE_Meilenstein" 'vorher "Zuordnung zu I-Stufe"
            self.info_fru_timming = DBInfo(
                path=self.config.get('default_path', '.'),
                attributes = FRUTimingAttribute,
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("FRU-Timing", strFRUTimingAttribute, rngFRUTimingAttribute, wbFRUTiming, wksFRUTiming, strFehlerFRUTiming, strDateinamen(5)) Then
#               blnImportAttribute(5) = True
            self.import_attribute[4] = self.info_fru_timming.einlesen_datei("FRU-Timing")
#           Else
            if not self.import_attribute[4]:
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeFRUTiming = ""
#               For i = LBound(strFRUTimingAttribute, 1) To UBound(strFRUTimingAttribute, 1)
#                   If strAttributeFRUTiming = "" Then
#                       strAttributeFRUTiming = strFRUTimingAttribute(i)
#                   Else
#                       strAttributeFRUTiming = strAttributeFRUTiming & ", " & strFRUTimingAttribute(i)
#                   End If
#               Next i
                #str_attribute_fru_timming = self.info_FRUTimming.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "FRU-Timing kann nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeFRUTiming & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "FRU-Timing kann nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeFRUTiming & ")"
#               End If
                self.errors += self.info_AVW.get_errors('FRU-Timing kann nicht eingelesen werden!')
#               blnImportAttribute(5) = False
#           End If
#       End If
#       
#       If blnAVWVorgaengerIDsVerwenden = True Then
        if self.use_predecessor_ids:
#           If blnImportAttribute(5) = True Then
            if self.import_attribute[4]:
#               'Arbeitsblatt AVWMaster_Rohdaten
#               'AVW: #1: ID, #2: temp1_Text, #3: Kommentar Redaktionskreis
#               'Weiche für Erfassung der Nachfolger-IDs
#               ReDim strAVWMasterAttribute(1 To 3)
#               ReDim rngAVWMandyAttribute(LBound(strAVWMasterAttribute, 1) To UBound(strAVWMasterAttribute, 1))
#               strAVWMasterAttribute(1) = "ID"
#               strAVWMasterAttribute(2) = "temp1_Text"
#               strAVWMasterAttribute(3) = "Kommentar Redaktionskreis"
                self.info_AVW_master = DBInfo(
                    path=self.config.get('default_path', '.'),
                    attributes = RequirementMasterAttribute
                )
#               'Dateiauswahl und Zuordnung
#               If EinlesenDatei("Anforderungen Masterbereich", strAVWMasterAttribute, rngAVWMasterAttribute, wbAVWMaster, wksAVWMaster, strFehlerAVWMaster, strDateinamen(6)) Then
#                   blnImportAttribute(6) = True
                self.import_attribute[5] = self.info_AVW_master.einlesen_datei("Anforderungen Masterbereich")
#               Else
                if not self.import_attribute[5]:
#                   'Sammlung aller gesuchten Attribute erzeugen
#                   strAttributeAVWMaster = ""
#                   For i = LBound(strAVWMasterAttribute, 1) To UBound(strAVWMasterAttribute, 1)
#                       If strAttributeAVWMaster = "" Then
#                           strAttributeAVWMaster = strAVWMasterAttribute(i)
#                       Else
#                           strAttributeAVWMaster = strAttributeAVWMaster & ", " & strAVWMasterAttribute(i)
#                       End If
#                   Next i
                    # strAttributeAVWMaster = self.info_AVW_master.str_attributes()
#                   If strFehlerGesamt = "" Then
#                       strFehlerGesamt = "Anforderungen aus dem Masterbereich können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVWMaster & ")"
#                   Else
#                       strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Anforderungen aus dem Masterbereich können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVWMaster & ")"
#                   End If
                    self.errors += self.info_AVW_master.get_errors('Anforderungen aus dem Masterbereich können nicht eingelesen werden!')
#                   blnImportAttribute(6) = False
#               End If
#           End If
#       End If
#       
#       'Rückgabewert
#       If blnAVWVorgaengerIDsVerwenden = False Then
#           If blnImportAttribute(1) And blnImportAttribute(2) And blnImportAttribute(3) And blnImportAttribute(4) And blnImportAttribute(5) Then
#               ATE_Status_Initializer = True
#           Else
#               ATE_Status_Initializer = False
#           End If
#       Else
#           If blnImportAttribute(1) And blnImportAttribute(2) And blnImportAttribute(3) And blnImportAttribute(4) And blnImportAttribute(5) And blnImportAttribute(6) Then
#               ATE_Status_Initializer = True
#           Else
#               ATE_Status_Initializer = False
#           End If
#       End If
        
        return all(self.import_attribute[:-1]) and (not self.use_predecessor_ids or self.import_attribute[5])
#   End Function

#
#   Private Sub ATE_Status_Deinitializer(ByRef wksBsM As Worksheet, ByRef wksAVW As Worksheet, ByRef wksTDVK As Worksheet, ByRef wksTDAA As Worksheet, ByRef wksTF As Worksheet, ByRef wksFRUTiming As Worksheet)
    def status_deinitialize(self):
        pass #TODO
#   'Worksheets identifizieren
#   Set wksBsM = Nothing
#   Set wksAVW = Nothing
#   Set wksTDVK = Nothing
#   Set wksTDAA = Nothing
#   Set wksTF = Nothing
#   Set wksFRUTiming = Nothing
#   Set verifikationKritList = Nothing
#   Set BsMDatenList = Nothing
#   Set testfallList = Nothing
#   Set FRUTimingList = Nothing
#   Set AVWVorgaengerList = Nothing
#   End Sub
#   
#    Funtion EinlesenDatei moved to om/db_info.py
#   
#    Function EinlesenDatei_Projecktspezifish moved to om/db_info.py
#   
#   Private Sub SchliessenWb(ByVal wbBsM As Workbook, ByVal wbAVW As Workbook, ByVal wbAVWMaster As Workbook, ByVal wbTDVK As Workbook, ByVal wbTDAA As Workbook, ByVal wbTF As Workbook, ByRef wbFRUTiming As Workbook)
    def close_workbooks(self) -> None:
        pass #TODO
#   'wbBsM.Close SaveChanges:=False
#   wbAVW.Close SaveChanges:=False
#   If Not wbAVWMaster Is Nothing Then
#       wbAVWMaster.Close SaveChanges:=False
#   End If
#   wbTDVK.Close SaveChanges:=False
#   wbTDAA.Close SaveChanges:=False
#   wbTF.Close SaveChanges:=False
#   wbFRUTiming.Close SaveChanges:=False
#   End Sub
#   
#   Private Function RangeObjekteVorhandenFehlerausgabe(ByRef rngObjekte() As Range, ByRef strObjekteNamen() As String, ByRef strNichtVorhandeneObjekte As String) As Boolean
    def RangeObjekteVorhandenFehlerausgabe(self) -> bool:
        pass #TODO
#       Dim i As Integer
#       
#       RangeObjekteVorhandenFehlerausgabe = True
#       strNichtVorhandeneObjekte = ""
#       For i = LBound(rngObjekte) To UBound(rngObjekte)
#           If rngObjekte(i) Is Nothing Then
#               RangeObjekteVorhandenFehlerausgabe = False
#               If strNichtVorhandeneObjekte = "" Then
#                   strNichtVorhandeneObjekte = strObjekteNamen(i)
#               Else
#                   strNichtVorhandeneObjekte = strNichtVorhandeneObjekte & ", " & strObjekteNamen(i)
#               End If
#           End If
#       Next i
#   End Function
#   
#   Private Sub EinlesenLAHBlacklist(ByVal wbBsM As Workbook, ByRef strLAHBlacklist() As String)
    def read_blacklist_LAHB(self):
#       Dim strWKSBlacklist As String           'String für Namen des Blacklist-Worksheets
#       Dim wksBlacklist As Worksheet           'Worksheet für Blacklist
#       Dim strAttributBlacklist As String      'String für Attribut der Blacklist
#       Dim rngBlacklist As Range               'Range für Blacklist
#       Dim lngBlacklist As Long                'Zählvariable für Blacklist
#       Dim lngBlacklistErfasst As Long         'Zählvariable für bereits erfasste Blacklist-Einträge
#       Dim blnBlacklistItemErfasst As Boolean  'Flag für bereits erfasste Blacklist-Items
#       Dim lngZeile As Long                    'Zeilenzähler
#       
#       On Error Resume Next
#       
#       strWKSBlacklist = "Blacklist"
        blacklist_name = self.config.get('blacklist_name', 'Blacklist')
#       strAttributBlacklist = "LAH, die ignoriert werden sollen"
        blacklist_attribute = self.config.get(
            'blacklist_attribute',
            'LAH, die ignoriert werden sollen'
        )
#       lngBlacklist = 0
        blacklist_len = 0
#       ReDim strLAHBlacklist(0)
        self.blacklist_LAHB = tuple()
#       
#       Set wksBlacklist = wbBsM.Sheets(strWKSBlacklist)
        blacklist_data_frame = self.workbook_BsM.sheet(blacklist_name)
#       If Not wksBlacklist Is Nothing Then
        if blacklist_data_frame is not None:
#           Set rngBlacklist = wksBlacklist.Cells.Find(strAttributBlacklist, lookat:=xlWhole)
#           If Not rngBlacklist Is Nothing Then
            if len(blacklist_data_frame) > 0 and blacklist_attribute in blacklist_data_frame.columns:
#               For lngZeile = 1 To wksBlacklist.UsedRange.Rows.Count - rngBlacklist.Row
                candidates = blacklist_data_frame[blacklist_attribute]
                self.blacklist_LAHB = tuple(
                    candidates[row]
                        for row in range(0, len(blacklist_data_frame))
                        if candidates[row] not in candidates[0:(row-1)]
                )               
                #for row in range(0, len(blacklist_data_frame)):
#                   blnBlacklistItemErfasst = False
                #    blacklist_item_reached = False
#                   If rngBlacklist.Offset(lngZeile, 0) <> "" Then
                #    if blacklist_data_frame[blacklist_attribute][row] != "":
#                       If lngBlacklist > 0 Then
                #        if blacklist_len > 0:
#                           For lngBlacklistErfasst = LBound(strLAHBlacklist, 1) To UBound(strLAHBlacklist, 1)
                #            for blacklist_item in blacklist_data_frame[blacklist_attribute]:
#                               If strLAHBlacklist(lngBlacklistErfasst) = rngBlacklist.Offset(lngZeile, 0) Then
                #                if self.blacklist_LAHB[-1] == blacklist_item:
#                                   blnBlacklistItemErfasst = True
                #                    blacklist_item_reached = True
#                                   Exit For
                #                    break
#                               End If
#                           Next lngBlacklistErfasst
#                           If blnBlacklistItemErfasst = False Then
                #            if blacklist_item_reached == False:
#                               lngBlacklist = lngBlacklist + 1
                #                blacklist_len += 1
#                               ReDim Preserve strLAHBlacklist(1 To lngBlacklist)
#                               strLAHBlacklist(lngBlacklist) = rngBlacklist.Offset(lngZeile, 0)
                #                self.blacklist_LAHB.append(blacklist_data_frame[blacklist_attribute][row])
#                           End If
#                       Else
#                           lngBlacklist = 1
                #            blacklist_len = 1
                
#                           ReDim strLAHBlacklist(1 To lngBlacklist)
#                           strLAHBlacklist(lngBlacklist) = rngBlacklist.Offset(lngZeile, 0)
                #            self.blacklist_LAHB.append(blacklist_data_frame[blacklist_attribute][row])
#                       End If
#                   End If
#               Next lngZeile

#           Else
            else:
#               MsgBox "Attribut der Blacklist """ & strAttributBlacklist & """ ist nicht vorhanden!"
                msgbox(f'Attribut der Blacklist "{blacklist_attribute}" is nicht vorhanden!')
#           End If
#       Else
        else:
#           MsgBox "Arbeitsblatt """ & strWKSBlacklist & """ ist nicht vorhanden!"
            msgbox(f'Arbeitsblatt "{blacklist_name}" is nicht vorhanden!')
#       End If
#   End Sub
#   
#   Private Sub EinlesenTDVKs(ByVal wksTDVK As Worksheet, ByRef strTDVKAttribute() As String, ByRef rngTDVKAttribute() As Range)
    def read_TDVKs(self):
#       Dim anfIDs As String                            'String für eingelesene Anforderungs-IDs
#       Dim verifikationKrit As Verifikationskriterium  'Klasse Verifikationskriterium
#       Dim idList As Collection                        'Liste mit Anforderungs-IDs, die zu einem Verifikationskriterium eingelesen werden
#       Dim lngZeile As Long                            'Long-Zähler für aktuell einzulesende Zeile
#       Dim strVerifikationsID As String                'String für ID des Verifikationskriteriums
#       
#       'Verifikationskriterien einlesen
#       'TDVKs: #1: ID, #2: Basierend auf der Anforderung, #3: Status, #4: Temp1_Text, #5: Aktion
#       For lngZeile = 1 To wksTDVK.UsedRange.Rows.Count - rngTDVKAttribute(1).Row
        for row in range(0, len(self.info_TDVK.columns)):
#       '    If rngTDVKAttribute(3).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Then
        #     if self.info_TDVK.columns[TDVCAttribute.Status][row] == 'Fachlich abgestimmt':
#               'Neues Verifikationskriterium anlegen
#               Set verifikationKrit = New Verifikationskriterium
            verification_criterion = Verificationskriterium(self.info_TDVK.columns, row)
                # initialization moved to Verificationskriterium.__init__
#               'Erfasstes Verifikationskriterium in globaler Verifikationskriterien-Liste hinzufügen
#               verifikationKritList.Add Item:=verifikationKrit, Key:=verifikationKrit.VK_ID
            self.verification_criteria[verification_criterion.vk_id] = verification_criterion
#       '    End If
#           'Fortschritt anzeigen
#           If lngZeile Mod 100 = 0 Then
                #TODO: Show progress
#               Debug.Print "Verifikationskriterien einlesen: " & lngZeile & "/" & wksTDVK.UsedRange.Rows.Count - rngTDVKAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub
#   
#   Private Sub EinlesenTDAAs(ByVal wksTDAA As Worksheet, ByRef strTDAAAttribute() As String, ByRef rngTDAAAttribute() As Range)
    def read_TDAAs(self):
        #Absicherungsauftrag (DE) == Security Order (EN)
#       Dim anfIDs As String                            'String für eingelesene Anforderungs-IDs
#       Dim absicherungsAuftr As Absicherungsauftraege  'Klasse Absicherungsauftraege
#       Dim lngZeile As Long                            'Long-Zähler für aktuell einzulesende Zeile
#       Dim strVerifikationsID As String                'String für ID des Verifikationskriteriums
#       Dim Verifikationskriterium As Verifikationskriterium    'Verifikationskriterium für Zuordnung
#       'Absicherungsaufträge einlesen
#       'TDAAs: #1: ID, #2: Enthalten in, #3: Status, #4: Testinstanz, #5: Testumgebungstyp
#       For lngZeile = 1 To wksTDAA.UsedRange.Rows.Count - rngTDAAAttribute(1).Row
        for row in range(0, len(self.info_TDAA.columns)):
#           If rngTDAAAttribute(3).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Or rngTDAAAttribute(3).Offset(lngZeile, 0).Value = "In Review" Or rngTDAAAttribute(3).Offset(lngZeile, 0).Value = "In Bearbeitung" Then
            if self.info_TDAA.columns[TDSafeGuardsAttribute.Status][row] in ['Fachlich abgestimmt', 'In Review', 'In Bearbeitung']:
                #  Security Order seems to be used only if the verification id is stored in self.verification_criterion_list,
                #  so we can check this before creating the security order
                # moved from above ------------------------
#               'ID des übergeordneten Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#               strVerifikationsID = Replace(Replace(rngTDAAAttribute(2).Offset(lngZeile, 0).Value, "?", ""), "r", "")
                verification_criterion_id = str(self.info_TDAA.columns[TDSafeGuardsAttribute.IncludedIn][row])
                verification_criterion_id = re.sub(r'[\?r]', '', verification_criterion_id)
                #-------------------------------------------
                verification_criterion = self.verification_criteria.get(verification_criterion_id, None)
#               'Neuen Absicherungsauftrag anlegen
#               Set absicherungsAuftr = New Absicherungsauftraege
                security_order:Absicherungsauftrag = Absicherungsauftrag(self.info_TDAA.columns, row)
                # initialization moved to Absicherungsauftrag.__init__

#               'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
#               Set Verifikationskriterium = New Verifikationskriterium
#               Set Verifikationskriterium = FindeVK(verifikationKritList, strVerifikationsID)
#               If Not Verifikationskriterium Is Nothing Then
#                   Verifikationskriterium.Absicherungsauftraege.Add Item:=absicherungsAuftr, Key:=absicherungsAuftr.abs_ID
#               End If
                if verification_criterion is not None:
                    verification_criterion.absicherungsauftraege[security_order.abs_id] = security_order
                    self.verification_criteria[verification_criterion_id] = verification_criterion
                ###TODO maybe None verification_criteria ids should be informed in a log file.
#           End If
#           'Fortschritt anzeigen
            ####TODO: Show progress
#           If lngZeile Mod 100 = 0 Then
#               Debug.Print "Absicherungsaufträge einlesen: " & lngZeile & "/" & wksTDAA.UsedRange.Rows.Count - rngTDAAAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub
#   
#   Private Sub EinlesenTFs(ByVal wksTF As Worksheet, ByRef strTFAttribute() As String, ByRef rngTFAttribute() As Range)
    def read_test_cases(self):
#       Dim lngZeile As Long                            'Long-Zähler für aktuell einzulesende Zeile
#       Dim testfall As Testfaelle                      'Klasse Testfaelle
#       Dim anfIDs As String                            'String für eingelesene Anforderungs-IDs
#       Dim idList As Collection                        'Liste mit Anforderungs-IDs
#       Dim varErfassteVKItem As Variant                'Variant für Item in der globalen Verifikationskriterien-Liste
#       Dim varAnfID As Variant                         'Anforderungs-ID aus varErfassteVKItem
#       Dim Verifikationskriterium As Verifikationskriterium    'Verifikationskriterium für Zuordnung
#       
#       'Testfälle einlesen
#       'TFs: #1: ID, #2: Status, #3: Testfallname, #4: Sonstige-Varianten, #5: Basierend auf Testdesign, #6: verifiziert, #7: Testinstanz
#       For lngZeile = 1 To wksTF.UsedRange.Rows.Count - rngTFAttribute(1).Row
        for row in range(0,len(self.info_TF.columns)):
#       '    If rngTFAttribute(2).Offset(lngZeile, 0).Value = "Operativ" Then
        #    if self.info_TF.columns[TestCaseAttribute.Status][row] == 'Operativ':
#               'Neuen Testfall anlegen
                # moved to TestCase.__init__
#               Set testfall = New Testfaelle
            test_case = TestCase(
                self.info_TF.columns,
                row,
                self.verification_criteria,
            )
#               
#               'Erfassten Testfall in globaler Testfall-Liste hinzufügen
#               testfallList.Add Item:=testfall, Key:=testfall.TF_ID
            self.test_cases[test_case.id] = test_case
#       '    End If
#           'Fortschritt anzeigen
            ###TODO: Show progress
#           If lngZeile Mod 100 = 0 Then
#               Debug.Print "Testfälle einlesen: " & lngZeile & "/" & wksTF.UsedRange.Rows.Count - rngTFAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub
#   
#   Private Sub EinlesenFRUTiming(ByVal wksFRUTiming As Worksheet, ByRef strFRUTimingAttribute() As String, ByRef rngFRUTimingAttribute() As Range)
    def read_FRU_timing(self):
#       Dim lngZeile As Long                    'Long-Zähler für aktuell einzulesende Zeile
#       Dim FRUTiming As FRUTiming              'Klasse FRUTiming
#       Dim strFRUKey As String                 'Key für Item in FRUTiming
#       
#       'Doppelte Einträge im FRU-Import ignorieren
#       On Error Resume Next
#       
#       'FRU-Timing eonlesen
#       'FRUTiming: #1: FeatureName, #2: RG, #3: Umsetzer, #4: Zuordnung zu I-Stufe
#       For lngZeile = 1 To wksFRUTiming.UsedRange.Rows.Count - rngFRUTimingAttribute(1).Row
        for row in range(0, len(self.info_fru_timming.columns)):
#           If rngFRUTimingAttribute(4).Offset(lngZeile, 0).Value <> "" Then
            if self.info_fru_timming.columns[FRUTimingAttribute.FEMilestone][row] != "":
#               'Neues FRUTiming anlegen
                ### moved to FRUTiming.__init__
#               Set FRUTiming = New FRUTiming
                fru_timing = FRUTiming(
                    self.info_fru_timming.columns,
                    row,
                    self.fru_timming_index
                )
#           End If
#           'Fortschritt anzeigen
#           If lngZeile Mod 100 = 0 Then
#               Debug.Print "FRU-Timing einlesen: " & lngZeile & "/" & wksFRUTiming.UsedRange.Rows.Count - rngFRUTimingAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub
#   
    def assign_verification_criterion(self, bsm_dataset:BSMData) -> None:
#       blnVKZugeordnet = False
#       For Each varErfassteVKItem In verifikationKritList
        for verification_criterion in self.verification_criteria.values():
#           'Abgleich über Element-ID
#           For Each varVKAnfID In varErfassteVKItem.anf_ids
            for requirement_id in verification_criterion.requirement_ids:
#               If varVKAnfID = BSMDatensatz.AVWID Then
                if  bsm_dataset.same_id(requirement_id):
#                   'Zugehörigkeit des Verifikationskriteriums zu aktuellen Anforderungen kennzeichnen
#                   varErfassteVKItem.AnforderungVorhanden = True
                    verification_criterion.requirement_present = True
#                   'Verifikationskriterium dem aktuellen AVW-Datensatz zuordnen
                    bsm_dataset.add_verification_criterion(verification_criterion)
                    ### moved to BSMData.add_verification_criterion()
#                   blnVKZugeordnet = True
#                   Exit For
                    return
#               End If
#           Next varVKAnfID
#           'Äußere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt
#           If blnVKZugeordnet Then Exit For
#       Next varErfassteVKItem

    def assign_test_cases(self, bsm_dataset:BSMData) -> None:
#       blnTFZugeordnet = False
        #test_case_assigned = False ###TODEL: Not used
#       For Each varErfassteTFItem In testfallList
        test_case:TestCase
        for test_case in self.test_cases.values():
#           'Abgleich über Element-ID
#           For Each varTFAnfID In varErfassteTFItem.TF_anfIDs
            for requirement_id in test_case.requirement_ids:
#               If varTFAnfID = BSMDatensatz.AVWID Then
                if bsm_dataset.same_id(requirement_id):
#                   BSMDatensatz.Testfaelle.Add Item:=varErfassteTFItem
                    bsm_dataset.test_cases.append(test_case)
#                   blnTFZugeordnet = True
                    #test_case_assigned = True ###TODEL: Not used
#                   Exit For
                    break
#               End If
#           Next varTFAnfID
#       Next varErfassteTFItem
#
    def add_dataset(self, row:int):
#       'Neuen AVW-Datensatz anlegen
        #Create new VW requirement dataset
        bsm_dataset = BSMData(
            self.info_AVW.columns,
            row,
            self.fru_timming_index,
            is_specific = self.is_project_specific,
        )
        ### Moved to BSMData.__init__
#       'Zusammenführung BsM-Relevanz
        bsm_dataset.set_relevance()
        ### refactored as BSMData.set_relevance()
        ### called from BSMData.__init__
#       'ASIL einlesen
        ### Moved to BSMData.__init__
#       
#       'Geplante I-Stufe einlesen
        ### refactored as BSMData.get_IStufe()
        ### called from BSMData.__init__
#       
        ### Moved to BSMData.__init__
#       
#       'Verifikationskriterium zuordnen
#       'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
        self.assign_verification_criterion(bsm_dataset)
#       
#       'Testfall zuordnen
        self.assign_test_cases(bsm_dataset)
#       
#       'Erfasste AVW-Rohdaten in globaler AVW-Rohdaten-Liste hinzufügen
#       BsMDatenList.Add Item:=BSMDatensatz, Key:=BSMDatensatz.AVWID
        self.bsm_datasets[bsm_dataset.avw_id] = bsm_dataset

    def add_successor_dataset(self, row:int):
#       'Neuen AVW-Datensatz anlegen
#       Set BSMDatensatz = New BSMDaten
        bsm_dataset = BSMSuccessorData(
            self.info_AVW.columns,
            row,
            self.fru_timming_index,
            self.use_predecessor_ids,
            self.predecessor_index_AVW,
            is_specific = self.is_project_specific,
        )
        ### Moved to BSMSuccessorData.__init__
#       'Zusammenführung BsM-Relevanz
        bsm_dataset.set_relevance()
        ### refactored as BSMData.set_relevance()
        ### called from BSMData.__init__
#       'ASIL einlesen
        ### Moved to BSMData.__init__
#       
#       'Geplante I-Stufe einlesen
        ### refactored as BSMData.get_IStufe()
        ### called from BSMData.__init__
#       
        ### Moved to BSMData.__init__
#       
#       'Verifikationskriterium zuordnen
#       'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
        self.assign_verification_criterion(bsm_dataset)
#       'Testfall zuordnen
        self.assign_test_cases(bsm_dataset)
#       
#       'Erfasste AVW-Rohdaten in globaler AVW-Rohdaten-Liste hinzufügen
#       BsMDatenList.Add Item:=BSMDatensatz, Key:=BSMDatensatz.AVWID
        self.bsm_datasets[bsm_dataset.avw_id] = bsm_dataset

#   Private Sub EinlesenAVWRohdaten(ByVal wksAVW As Worksheet, ByRef strAVWAttribute() As String, ByRef rngAVWAttribute() As Range, ByRef strLAHBlacklist() As String, _
#                                   ByVal strProjekt As String, ByRef strAVWAttributeMEB21() As String, ByRef rngAVWAttributeMEB21() As Range)
    def read_requirement_raw_data(self):
#       Dim lngZeile As Long                    'Long-Variable für aktuell einzulesende Zeile
#       Dim BSMDatensatz As BSMDaten            'Klasse BSMDaten
#       Dim strBsMRelevanz As String            'String-Variable für Zusammenfassung der BsM-Relevanz
#       Const strBsMVorhanden As String = "ja"  'Konstante für Angabe aus AVW für vorhandenes BsM-Attribut
#       Dim varErfassteVKItem As Variant        'Variant für Item in der globalen Verifikationskriterien-Liste
#       Dim varVKAnfID As Variant               'Anforderungs-ID aus varErfassteVKItem
#       Dim blnVKZugeordnet As Boolean          'Flag für zugeordnetes Verifikationskriterium
#       Dim varErfassteTFItem As Variant        'Variant für Item in der globalen Testfälle-Liste
#       Dim varTFAnfID As Variant               'Anforderungs-ID aus varErfassteTFItem
#       Dim blnTFZugeordnet As Boolean          'Flag für zugeorndeten Testfall
#       Dim varUmsetzer As Variant              'Variant-Array für Zerlegung der Umsetzer
#       Dim intUmsetzer As Integer              'Zählvariable für Umsetzer
#       Dim strIStufe As String                 'String für gefundene I-Stufe
#       Dim strIStufeMin As String              'String für früheste I-Stufe
#       Dim lngLAHBlacklist As Long             'Long für Zähler der Blacklist
#       
#       'Fehlerbehandlung ausschalten für evtl. fehlende Keys in der Collection FRUTimingList
#       On Error Resume Next
#       
#       'Anforderungen einlesen
#       'AVW: #1: ID, #2: Dokument-ID, #3: Basis für Testdesign, #4: Typ, #5: Kategorie, #6: Status, #7: Feature, #8: Reifegrad, #9: Umsetzer, #10: ASIL
#       '#11: BSM-SaFuSi Bewertung, #12: BSM-ZZ Bewertung, #13: BSM-ED Bewertung, #14: BSM-FFF Bewertung, #15: BSM-O Bewertung, #16: BSM-Se Bewertung, #17: MV
#       '#18: Cluster Testing, #19: Dokument, #20: Kommentar Redaktionskreis, #21: temp1_Text
#       For lngZeile = 1 To wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
        for row in range(0, len(self.info_AVW.columns)):
#           'Nur Datensätze bei vorhandener Anforderungs-ID übernehmen
#           If rngAVWAttribute(1).Offset(lngZeile, 0).Value <> "" Then
            if str(self.info_AVW.columns[RequirementAttribute.ID][row]) != "":
#               'Nur LAH aufnehmen, die nicht auf der Blacklist stehen
#               If AuswertungLAHBlacklist(strLAHBlacklist, rngAVWAttribute(19).Offset(lngZeile, 0).Value) Then
                if str(self.info_AVW.columns[RequirementAttribute.Document][row]) not in self.blacklist_LAHB:
#                   'Nur Fachlich abgestimmte (AVW) oder gültige (DOORS) Anforderungen aufnehmen
                    #Only include requirements that are either 
                    # "Fachlich abgestimmt" -- Technically agreed requirements ;or 
                    # "gültig"              -- valid DOORS requirements. 
                    # The status is stored in column 6, which is accessed via self.info_AVW.columns[RequirementAttribute.Status][row]
#           '        If rngAVWAttribute(6).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Or rngAVWAttribute(6).Offset(lngZeile, 0).Value = "gültig" Then
            #        if self.info_AVW.columns[RequirementAttribute.Status][row] in ['Fachlich abgestimmt', 'gültig']:
#                       'Attribut Cluster Testing => nicht relevant soll ignoriert werden
                        # Testing cluster attribute => "nich relevant" not relevant should be ignored
#           '            If rngAVWAttribute(18).Offset(lngZeile, 0).Value <> "nicht relevant" Then
            #            if str(self.info_AVW.columns[RequirementAttribute.TestingCluster][row]) != "nicht relevant":
                    self.add_dataset(row)
#           '            End If
#           '        End If
#               End If
#           End If
#       
#           'Fortschritt anzeigen
            #TODO: Show progress
#           If lngZeile Mod 100 = 0 Then
#               Debug.Print "Anforderungen einlesen: " & lngZeile & "/" & wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub

#   Private Sub EinlesenAVWNachfolgerRohdaten(ByVal wksAVW As Worksheet, ByRef strAVWAttribute() As String, ByRef rngAVWAttribute() As Range, ByRef strLAHBlacklist() As String, _
#                                             ByVal strProjekt As String, ByRef strAVWAttributMEB21() As String, ByRef rngAVWAttributMEB21() As Range)
    def read_successor_requirement_raw_data(self):
        pass #TODO
#       Dim lngZeile As Long                    'Long-Variable für aktuell einzulesende Zeile
#       Dim BSMDatensatz As BSMDaten            'Klasse BSMDaten
#       Dim strBsMRelevanz As String            'String-Variable für Zusammenfassung der BsM-Relevanz
#       Const strBsMVorhanden As String = "ja"  'Konstante für Angabe aus AVW für vorhandenes BsM-Attribut
#       Dim varErfassteVKItem As Variant        'Variant für Item in der globalen Verifikationskriterien-Liste
#       Dim varVKAnfID As Variant               'Anforderungs-ID aus varErfassteVKItem
#       Dim blnVKZugeordnet As Boolean          'Flag für zugeordnetes Verifikationskriterium
#       Dim varErfassteTFItem As Variant        'Variant für Item in der globalen Testfälle-Liste
#       Dim varTFAnfID As Variant               'Anforderungs-ID aus varErfassteTFItem
#       Dim blnTFZugeordnet As Boolean          'Flag für zugeorndeten Testfall
#       Dim varUmsetzer As Variant              'Variant-Array für Zerlegung der Umsetzer
#       Dim intUmsetzer As Integer              'Zählvariable für Umsetzer
#       Dim strIStufe As String                 'String für gefundene I-Stufe
#       Dim strIStufeMin As String              'String für früheste I-Stufe
#       Dim lngLAHBlacklist As Long             'Long für Zähler der Blacklist
#       Dim AVWVorgaenger As AVWVorgaenger      'Klasse AVWVorgaenger
#       
#       'Fehlerbehandlung ausschalten für evtl. fehlende Keys in der Collection FRUTimingList
#       On Error Resume Next
#       
#       'Anforderungen einlesen
#       'AVW: #1: ID, #2: Dokument-ID, #3: Basis für Testdesign, #4: Typ, #5: Kategorie, #6: Status, #7: Feature, #8: Reifegrad, #9: Umsetzer, #10: ASIL
#       '#11: BSM-SaFuSi Bewertung, #12: BSM-ZZ Bewertung, #13: BSM-ED Bewertung, #14: BSM-FFF Bewertung, #15: BSM-O Bewertung, #16: BSM-Se Bewertung, #17: MV
#       '#18: Cluster Testing, #19: Dokument, #20: Kommentar Redaktionskreis, #21: temp1_Text, #22: ID der Vorgänger-Anforderung
#       For lngZeile = 1 To wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
        for row in range(0, len(self.info_AVW.columns)):
#           'Nur Datensätze bei vorhandener Anforderungs-ID übernehmen
#           If rngAVWAttribute(1).Offset(lngZeile, 0).Value <> "" Then
            if str(self.info_AVW.columns[RequirementAttribute.ID][row]) != "":
#               'Nur LAH aufnehmen, die nicht auf der Blacklist stehen
#               If AuswertungLAHBlacklist(strLAHBlacklist, rngAVWAttribute(19).Offset(lngZeile, 0).Value) Then
                if str(self.info_AVW.columns[RequirementAttribute.Document][row]) not in self.blacklist_LAHB:
#                   'Nur Fachlich abgestimmte (AVW) oder gültige (DOORS) Anforderungen aufnehmen
#           '        If rngAVWAttribute(6).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Or rngAVWAttribute(6).Offset(lngZeile, 0).Value = "gültig" Then
            #        if self.info_AVW.columns[RequirementAttribute.Status][row] in ['Fachlich abgestimmt', 'gültig']:
#                       'Attribut Cluster Testing => nicht relevant soll ignoriert werden
#           '            If rngAVWAttribute(18).Offset(lngZeile, 0).Value <> "nicht relevant" Then
            #            if str(self.info_AVW.columns[RequirementAttribute.TestingCluster][row]) != "nicht relevant":
                    # refactored in add_dataset
                    self.add_successor_dataset(row)
#           '            End If
#           '        End If
#               End If
#           End If
#       
#           'Fortschritt anzeigen
            ###TODO: Show progress
#           If lngZeile Mod 100 = 0 Then
#               Debug.Print "Anforderungen einlesen: " & lngZeile & "/" & wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub
#   
#   Private Sub EinlesenAVWVorgaengerRohdaten(ByVal wksAVWMaster As Worksheet, ByRef strAVWMasterAttribute() As String, ByRef rngAVWMasterAttribute() As Range)
    def read_predecessor_requirement_raw_data(self):
#       Dim lngZeile As Long                    'Long-Variable für aktuell einzulesende Zeile
#       Dim AVWVorgaenger As AVWVorgaenger      'Klasse AVWVorgaenger
#       
#       'Anforderungen einlesen
#       'AVW: #1: ID, #2: temp1_Text, #3: Kommentar Redaktionskreis
#       For lngZeile = 1 To wksAVWMaster.UsedRange.Rows.Count - rngAVWMasterAttribute(1).Row
        for row in range(0, len(self.info_AVW_master.columns)):
#           'Nur Datensätze bei vorhandener Anforderungs-ID übernehmen
#           If rngAVWMasterAttribute(1).Offset(lngZeile, 0).Value <> "" Then
            if str(self.info_AVW_master.columns[RequirementMasterAttribute.ID][row]) != "":
#               'Neuen AVW-Datensatz anlegen
#               Set AVWVorgaenger = New AVWVorgaenger
#               'Master-ID einlesen
                ### Moved to AVWVorgaenger.__init__
                vw_requirement_predecessor = VWRequirementPredecessor(
                    self.info_AVW_master.columns,
                    row,
                )
#               
#               'Erfasste AVW-Vorgänger in globaler AVW-Vorgaenger-Liste hinzufügen
#               AVWVorgaengerList.Add Item:=AVWVorgaenger, Key:=AVWVorgaenger.ID
                self.predecessor_index_AVW[vw_requirement_predecessor.id] = vw_requirement_predecessor
#           End If
#       
#           'Fortschritt anzeigen
            ###TODO: Show progress
#           If lngZeile Mod 100 = 0 Then
#               Debug.Print "Anforderungen aus Masterbereich einlesen: " & lngZeile & "/" & wksAVWMaster.UsedRange.Rows.Count - rngAVWMasterAttribute(1).Row
#           End If
#       Next lngZeile
#   End Sub
#   
#   Function EinlesenGetrennteWerteKomma(ByVal lahIDs As String) As Collection
    #Moved to utils as it could be used in multiple places.
#       
#       Private Function AusgabeSammlungLF(ByRef list As Collection) As String
#       Dim strTemp As String
#       Dim i As Integer
#       
#       strTemp = ""
#       If list.Count > 0 Then
#           For i = 1 To list.Count
#               If (strTemp = "") Then
#                   strTemp = list(i)
#               Else
#                   strTemp = strTemp & vbCrLf & list(i)
#               End If
#           Next
#       End If
#       AusgabeSammlungLF = strTemp
#   End Function
#   
#   Private Function AusgabeSammlungLFEinfach(ByRef list As Collection) As String
    def AusgabeSammlungLFEinfach(self, input_list:list[str]|tuple[str]):
        ###TOBEDEL: it can be replaced by a join sentence.
        return unic_join(CRLF, input_list)
#       Dim strTemp As String
#       Dim i As Integer
#       
#       strTemp = ""
#       If list.Count > 0 Then
#           For i = 1 To list.Count
#               If (strTemp = "") Then
#                   strTemp = list(i)
#               Else
#                   If InStr(strTemp, list(i)) = 0 Then
#                       strTemp = strTemp & vbCrLf & list(i)
#                   End If
#               End If
#           Next
#       End If
#       AusgabeSammlungLFEinfach = strTemp
#   End Function
#   
#   Private Function AusgabeSammlungKomma(ByRef list As Collection) As String
    def AusgabeSammlungKomma(self):
        ###TOBEDEL: it can be replaced by a join sentence.
        return ', '.join(list)
#       Dim strTemp As String
#       Dim i As Integer
#       
#       strTemp = ""
#       If list.Count > 0 Then
#           For i = 1 To list.Count
#               If (strTemp = "") Then
#                   strTemp = list(i)
#               Else
#                   strTemp = strTemp & ", " & list(i)
#               End If
#           Next
#       End If
#       AusgabeSammlungKomma = strTemp
#   End Function
#   
#   Private Function AuswertungLAHBlacklist(ByRef strLAHBlacklist() As String, ByVal strLAH As String)
    def blacklist_evaluation_LAH(self, str_lah:str) -> bool:
        ### TOBEDEL Unused. Non required use "in" operator instead.
#       Dim lngBlacklistZaehler As Long     'Long-Zähler für Blacklist
#       
#       AuswertungLAHBlacklist = True
#       If UBound(strLAHBlacklist, 1) > 0 Then
#           For lngBlacklistZaehler = LBound(strLAHBlacklist, 1) To UBound(strLAHBlacklist, 1)
#               If strLAHBlacklist(lngBlacklistZaehler) = strLAH Then
#                   AuswertungLAHBlacklist = False
#                   Exit For
#               End If
#           Next lngBlacklistZaehler
#       End If
        return not str_lah in self.blacklist_LAHB
#   End Function
#   
#   Private Function FindeVK(ByRef Liste As Collection, ByVal strKey As String) As Verifikationskriterium
    def FindeVK(self, liste:dict[str,Verificationskriterium], key:str) -> Verificationskriterium:
        return liste.get(key,None)
        #TOBEDEL Non required use get method instead.
#       Dim ListenObjekt As Verifikationskriterium
#           
#       On Error GoTo err
#       Set ListenObjekt = Liste.Item(strKey)
#       Set FindeVK = ListenObjekt
#       Exit Function
#       err:
#           Set ListenObjekt = Nothing
#           Set FindeVK = ListenObjekt
#   End Function
#   
#   Private Function FindeAVWVorgaenger(ByRef Liste As Collection, ByVal strKey As String) As AVWVorgaenger
    def find_predecessor_AVW(self, values:dict[str,VWRequirementPredecessor], key:str) -> VWRequirementPredecessor|None:
        return values.get(key, None)
        #TOBEDEL Unused. Non required use get method instead.
#       Dim ListenObjekt As AVWVorgaenger
#           
#       On Error GoTo err
#       Set ListenObjekt = Liste.Item(strKey)
#       Set FindeAVWVorgaenger = ListenObjekt
#       Exit Function
#       err:
#           Set ListenObjekt = Nothing
#           Set FindeAVWVorgaenger = ListenObjekt
#   End Function

    def handle_test_case(self, tc:TestCase):
        if tc.test_environment_type in self.relevant_test_environments:
            # Manage relevant test environment test case
            test_environment_index = self.relevant_test_environments.index(tc.test_environment_type)
            if tc.status == "Operativ":
                if self.te_comparison_count[test_environment_index] in (0,20):
                    self.te_comparison_count[test_environment_index] += 10
            else:
                if self.te_comparison_count[test_environment_index] in (0,10):
                    self.te_comparison_count[test_environment_index] += 20
            # Manage other test environment test cases
        elif  tc.test_environment_type not in self.other_test_environment:
            self.other_test_environment.append(tc.test_environment_type)
#               blnTFTUZugeordnet = False
#               'Erfassung der vorhandenen relevanten Testumgebungen
#               For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
#                   If varErfassteTFItem.TF_Testumgebungstyp = strAbgleichTUs(i) Then
#                       blnTFTUZugeordnet = True
#                       'Unterscheidung nach Status des Testfalls
#                       If varErfassteTFItem.TF_Status = "Operativ" Then
#                           'Nicht operative Testfälle bereits erfasst?
#                           If intAbgleichTUs(i) = 0 Then
#                               intAbgleichTUs(i) = 10
#                           ElseIf intAbgleichTUs(i) = 20 Then
#                               intAbgleichTUs(i) = 30
#                           End If
#                       Else
#                           'Operative Testfälle bereits erfasst?
#                           If intAbgleichTUs(i) = 0 Then
#                               intAbgleichTUs(i) = 20
#                           ElseIf intAbgleichTUs(i) = 10 Then
#                               intAbgleichTUs(i) = 30
#                           End If
#                       End If
#                   End If
#               Next i
#   '           'Restliche bekannte TUs abgleichen
#   '           If blnTFTUZugeordnet = False Then
#   '               For i = intRelevantekTUs + 1 To UBound(strBekannteTUs, 1)
#   '                    If varErfassteTFItem.TF_Testumgebungstyp = strBekannteTUs(i) Then
#   '                        blnTFTUZugeordnet = True
#   '                        Exit For
#   '                    End If
#   '                Next i
#   '            End If
#               'Weitere TUs erfassen
#               If blnTFTUZugeordnet = False Then
#                   If intWeitereTUs > 0 Then
                    #if len(self.other_test_environment) > 0:
#                       For intWeitereTUsZaehler = 1 To intWeitereTUs
#                           If strWeitereTUs(intWeitereTUsZaehler) = varErfassteTFItem.TF_Testumgebungstyp Then
#                               blnTFTUZugeordnet = True
#                               Exit For
#                           End If
#                       Next intWeitereTUsZaehler
#                       If blnTFTUZugeordnet = False Then
#                           intWeitereTUs = intWeitereTUs + 1
#                           ReDim Preserve strWeitereTUs(1 To intWeitereTUs)
#                           strWeitereTUs(intWeitereTUs) = varErfassteTFItem.TF_Testumgebungstyp
#                       End If
#                   Else
#                       intWeitereTUs = 1
#                       ReDim strWeitereTUs(1 To intWeitereTUs)
#                       strWeitereTUs(intWeitereTUs) = varErfassteTFItem.TF_Testumgebungstyp
#                   End If
#               End If
#   
    def auswertung_tf(self, test_cases:list[TestCase]):
#       'Auswertung TF
#       strTestfaelle = ""
        str_test_cases = ''
#       If varErfassteBsMDatensatzItem.Testfaelle.Count > 0 Then
        if len(test_cases) > 0:
            str_test_cases = CRLF.join([f"{tc.id} - {tc.status} - {tc.test_instance} - {tc.test_environment_type}" for tc in test_cases]) 
#           For Each varErfassteTFItem In varErfassteBsMDatensatzItem.Testfaelle
            tc:TestCase
            for tc in test_cases:
                #---------------- moved above outside the loop
#               'Testfälle zusammenführen
#               If strTestfaelle = "" Then
#                   strTestfaelle = varErfassteTFItem.TF_ID & " - " & varErfassteTFItem.TF_Status & " - " & varErfassteTFItem.TF_Testinstanz & " - " & varErfassteTFItem.TF_Testumgebungstyp
#               Else
#                   strTestfaelle = strTestfaelle & vbCrLf & varErfassteTFItem.TF_ID & " - " & varErfassteTFItem.TF_Status & " - " & varErfassteTFItem.TF_Testinstanz & " - " & varErfassteTFItem.TF_Testumgebungstyp
#               End If
#   
                #----------------
                self.handle_test_case(tc)
                # Moved to handle_test_case method
#           Next varErfassteTFItem
#       End If
        return str_test_cases

    def abgleich_TUs(self, security_order:Absicherungsauftrag):
#       'Abgleich der vorhandenen relevanten Testumgebungen
#       blnAATUZugeordnet = False
        blnAATUZugeordnet = False
#       For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
        for index, test_environment_comparison in enumerate(self.relevant_test_environments):
#           If varErfassteTDAAItem.Testumgebungstyp = strAbgleichTUs(i) Then
            if security_order.test_environment_type == test_environment_comparison:
#               blnAATUZugeordnet = True
                blnAATUZugeordnet = True
#               If intAbgleichTUs(i) = 0 Then
#                   'VK-TU vorhanden, kein Testfall vorhanden
#                   intAbgleichTUs(i) = 1
#               ElseIf intAbgleichTUs(i) = 10 Then
#                   'VK-TU vorhanden, Testfälle operativ
#                   intAbgleichTUs(i) = 11
#               ElseIf intAbgleichTUs(i) = 20 Then
#                   'VK-TU vorhanden, Testfälle nicht operativ
#                   intAbgleichTUs(i) = 21
#               ElseIf intAbgleichTUs(i) = 30 Then
#                   'VK-TU vorhanden, Testfälle operativ und nicht operativ
#                   intAbgleichTUs(i) = 31
                if self.te_comparison_count[index] in (0, 10, 20, 30):
                    self.te_comparison_count[index] += 1
#               End If
                break # optimization: only a test_environment can match
#           End If
#       Next i
#'       'Restliche bekannte TUs abgleichen
#'       If blnAATUZugeordnet = False Then
#'           For i = intRelevantekTUs + 1 To UBound(strBekannteTUs, 1)
#'               If varErfassteTDAAItem.Testumgebungstyp = strBekannteTUs(i) Then
#'                   blnAATUZugeordnet = True
#'                   Exit For
#'               End If
#'           Next i
#'       End If
#       'Weitere TUs erfassen
        #------------------------------------------------------
#       If blnAATUZugeordnet = False Then
#           If intWeitereTUs > 0 Then
#               For intWeitereTUsZaehler = 1 To intWeitereTUs
#                   If strWeitereTUs(intWeitereTUsZaehler) = varErfassteTDAAItem.Testumgebungstyp Then
#                       blnAATUZugeordnet = True
#                       Exit For
#                   End If
#               Next intWeitereTUsZaehler
#               If blnAATUZugeordnet = False Then
#                   intWeitereTUs = intWeitereTUs + 1
#                   ReDim Preserve strWeitereTUs(1 To intWeitereTUs)
#                   strWeitereTUs(intWeitereTUs) = varErfassteTDAAItem.Testumgebungstyp
#               End If
#           Else
#               intWeitereTUs = 1
#               ReDim strWeitereTUs(1 To intWeitereTUs)
#               strWeitereTUs(intWeitereTUs) = varErfassteTDAAItem.Testumgebungstyp
#           End If
#       End If
        if not blnAATUZugeordnet:
            if security_order.test_environment_type not  in self.other_test_environment:
                self.other_test_environment.append(security_order.test_environment_type)

#   Private Sub AusgabeATEStatus(ByVal wbBsM As Workbook, ByRef wksBsM As Worksheet, ByRef strBsMAttribute() As String, ByRef rngBsMAttribute() As Range, ByRef strWeitereTUsAusgabe As String, ByRef strDateinamen() As String, ByVal strProjekt As String)
    def output_status(self) -> tuple[dict[str,list[str]],str]:
#       Dim lngDatensatz As Long                        'Long-Variable für aktuell zu schreibenden Datensatz
#       Dim varErfassteBsMDatensatzItem As Variant      'Variant für Item im globalen BsM-Datensatz
#       Dim varErfassteTDAAItem As Variant              'Variant für Item aus den jeweiligen Absicherungsaufträgen
#       Dim strTDAA As String                           'String für Sammlung der Absicherungsaufträge
#       Dim strTDTiTu As String                         'String für Sammlung der Ti:Tu-Kombinationen
#       Dim varErfassteTFItem As Variant                'Variant für Item aus den jeweiligen Testfällen
#       Dim strTestfaelle As String                     'String für Sammlung der Testfälle
#       Dim strAbgleichTUs() As String                  'String-Array für die Namen der abzugleichenden Testumgebungen bei TDs und TFs
#       Dim intAbgleichTUs() As Integer                 'Integer-Array für Erfassung der Testumgebungen bei TDs und TFs
#       Dim i As Long                                   'Laufvariable
#       Dim intAuswertungTUs() As Integer               'Integer-Array für die Ergebnisse des Tu-Abgleichs
#       Dim strAuswertungTUs() As String                'String-Array für die Ergebnisse des Tu-Abgleichs
#       Dim strAusgabeAuswertungTUs As String           'String für Ausgabe des Tu-Abgleichs
#       Dim intAusgabeAuswertungTUs As Integer          'Integer für Ausgabe des Tu-Abgleichs
#       Dim strAuswertungTUsFehlendeAAs As String       'String für Ausgabe der fehlenden TUs bei TD-AAs
#       Dim strAuswertungTUsFehlendeTFs As String       'String für Ausgabe der fehlenden TUs bei TFs
#       Dim strAusgabeAuswertungTUsDetails As String    'String für Ausgabe des Tu-Abgleichs mit Details
#       Dim intWeitereTUs As Integer                    'Zählvariable für weitere TUs
#       Dim strWeitereTUs() As String                   'String-Array für weitere TUs
        self.other_test_environment = []
#       Dim strBekannteTUs() As String                  'String-Array für alle bekannten TUs
#       Dim blnTFTUZugeordnet As Boolean                'Flag für zugeordnete TU
#       Dim blnAATUZugeordnet As Boolean                'Flag für zugeordnete TU
#       Dim intWeitereTUsZaehler As Integer             'Laufvariable für weitere TUs
#       Dim blnTIUnerlaubt As Boolean                   'Flag für unerlaubte Testinstanzen (Erlaubt: "eigene Organisationseinheit", "Dauerlauf Gesamtfahrzeug", "Gesamtverbundintegration" sowie alle eingetragenen Umsetzer)
#       Dim intUmsetzer As Integer                      'Zähler für Umsetzer aus AVW
#       Dim varUmsetzer As Variant                      'Variant-Array für Umsetzer aus AVW
#       Dim blnUmsetzer() As Boolean                    'Flag-Array für Abgleich der Umsetzer AVW<->AA
#       Dim intZielspalte As Integer                    'Integer für Ausgabespalte
#       Dim intRelevantekTUs As Integer                 'Integer für Anzahl der relevanten TUs
#       Dim dblTDVKAnzahlUseCases As Double             'Anzahl der Vorkommen der Use-Case-Begriffe
#       Dim strTDVKAktion As String                     'String zur Bearbeitung der TDVK-Aktion
#       
        bg_colours:FillDict = FillDict()
#       'Tabelle erzeugen
#       'Neues Worksheet erzeugen
#       Set wksBsM = wbBsM.Sheets.Add(after:=wbBsM.Worksheets(wbBsM.Worksheets.Count))
#       wksBsM.Name = "ATE_Status_" & "Today" & "_" & Replace(Time, ":", "")
        worksheet_name = f'ATE_Status_{self.date_suffix}'
#       'Arbeitsblatt BsM_Status
#       'BsM-Status: #1: ID, #2: Dokument-ID, #3: BsM-Relevanz, #4: BSM-SaFuSi Bewertung, #5: BSM-ZZ Bewertung, #6: BSM-ED Bewertung, #7: BSM-FFF Bewertung, #8: BSM-O Bewertung,
#       '#9: BSM-Se Bewertung, #10: ASIL, #11: Feature, #12: Reifegrad, #13: Umsetzer, #14: Status, #15: TD-VK, #16: TD-AA, #17: TD-TI:TU, #18: Testfälle, #19: Vergleich TUs,
#       '#20: MV, #21: Kategorie, #22: Dokument, #23: #abgelehnt_nicht_testbar, #24: Zugeordnete I-Stufe, #25: Status TD-VK, #26: Fehlende TUs bei TD-AAs, #27: Fehlende TUs bei TFs,
#       '#28: Erläuterungen zum Vergleich, #29: Cluster Testing, #30: Projekt, #31: TD-VK temp1_Text, #32: TD-VK Effort Estimation, #33: Anforderungsverantwortliche,#34: KW
#       'Projektspezifisch - MEB21
#       '#35: Temp11_Auswahlfeld
#       
#       lngDatensatz = 1
        bsm_attributes = [field.value for field in OutputBSMAttribute]
#       If blnAVWVorgaengerIDsVerwenden = False Then
#           intZielspalte = 0
#           ReDim strBsMAttribute(1 To 34)
#       Else
        if not self.use_predecessor_ids:
#           intZielspalte = 1
#           ReDim strBsMAttribute(0 To 34)
            bsm_attributes.remove(OutputBSMAttribute.RedirectedFrom)
#       End If
        
        ###Moved below to combine with is_project_specific condition
#       ReDim rngBsMAttribute(LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1))
#       'Name und Position der Tabellenattribute
#       If blnAVWVorgaengerIDsVerwenden Then
#           strBsMAttribute(0) = "Abgezweigt aus"
#           Set rngBsMAttribute(0) = wksBsM.Cells(lngDatensatz, intZielspalte)
#       End If
#       strBsMAttribute(34) = "KW Datenauswertung"
#       Set rngBsMAttribute(34) = wksBsM.Cells(lngDatensatz, intZielspalte + 1)
#       strBsMAttribute(1) = "ID"
#       Set rngBsMAttribute(1) = wksBsM.Cells(lngDatensatz, intZielspalte + 2)
#       strBsMAttribute(2) = "Dokument-ID"
#       Set rngBsMAttribute(2) = wksBsM.Cells(lngDatensatz, intZielspalte + 3)
#       strBsMAttribute(22) = "Dokument"
#       Set rngBsMAttribute(22) = wksBsM.Cells(lngDatensatz, intZielspalte + 4)
#       strBsMAttribute(21) = "Kategorie"
#       Set rngBsMAttribute(21) = wksBsM.Cells(lngDatensatz, intZielspalte + 5)
#       strBsMAttribute(11) = "Feature"
#       Set rngBsMAttribute(11) = wksBsM.Cells(lngDatensatz, intZielspalte + 6)
#       strBsMAttribute(12) = "Reifegrad"
#       Set rngBsMAttribute(12) = wksBsM.Cells(lngDatensatz, intZielspalte + 7)
#       strBsMAttribute(13) = "Umsetzer"
#       Set rngBsMAttribute(13) = wksBsM.Cells(lngDatensatz, intZielspalte + 8)
#       strBsMAttribute(3) = "BsM-Relevanz"
#       Set rngBsMAttribute(3) = wksBsM.Cells(lngDatensatz, intZielspalte + 9)
#       strBsMAttribute(4) = "BSM-SaFuSi Bewertung"
#       Set rngBsMAttribute(4) = wksBsM.Cells(lngDatensatz, intZielspalte + 10)
#       strBsMAttribute(5) = "BSM-ZZ Bewertung"
#       Set rngBsMAttribute(5) = wksBsM.Cells(lngDatensatz, intZielspalte + 11)
#       strBsMAttribute(6) = "BSM-ED Bewertung"
#       Set rngBsMAttribute(6) = wksBsM.Cells(lngDatensatz, intZielspalte + 12)
#       strBsMAttribute(7) = "BSM-FFF Bewertung"
#       Set rngBsMAttribute(7) = wksBsM.Cells(lngDatensatz, intZielspalte + 13)
#       strBsMAttribute(8) = "BSM-O Bewertung"
#       Set rngBsMAttribute(8) = wksBsM.Cells(lngDatensatz, intZielspalte + 14)
#       strBsMAttribute(9) = "BSM-Se Bewertung"
#       Set rngBsMAttribute(9) = wksBsM.Cells(lngDatensatz, intZielspalte + 15)
#       strBsMAttribute(10) = "ASIL"
#       Set rngBsMAttribute(10) = wksBsM.Cells(lngDatensatz, intZielspalte + 16)
#       strBsMAttribute(14) = "Status"
#       Set rngBsMAttribute(14) = wksBsM.Cells(lngDatensatz, intZielspalte + 17)
#       strBsMAttribute(29) = "Cluster Testing"
#       Set rngBsMAttribute(29) = wksBsM.Cells(lngDatensatz, intZielspalte + 18)
#       strBsMAttribute(23) = "#abgelehnt_nicht_testbar"
#       Set rngBsMAttribute(23) = wksBsM.Cells(lngDatensatz, intZielspalte + 19)
#       strBsMAttribute(20) = "MV"
#       Set rngBsMAttribute(20) = wksBsM.Cells(lngDatensatz, intZielspalte + 20)
#       strBsMAttribute(33) = "Anforderungsverantwortlicher"
#       Set rngBsMAttribute(33) = wksBsM.Cells(lngDatensatz, intZielspalte + 21)
#       strBsMAttribute(15) = "TD-VK"
#       Set rngBsMAttribute(15) = wksBsM.Cells(lngDatensatz, intZielspalte + 22)
#       strBsMAttribute(25) = "Status TD-VK"
#       Set rngBsMAttribute(25) = wksBsM.Cells(lngDatensatz, intZielspalte + 23)
#       strBsMAttribute(31) = "TD-VK temp1_Text"
#       Set rngBsMAttribute(31) = wksBsM.Cells(lngDatensatz, intZielspalte + 24)
#       strBsMAttribute(32) = "TD-VK Effort Estimation"
#       Set rngBsMAttribute(32) = wksBsM.Cells(lngDatensatz, intZielspalte + 25)
#       strBsMAttribute(16) = "TD-AA"
#       Set rngBsMAttribute(16) = wksBsM.Cells(lngDatensatz, intZielspalte + 26)
#       strBsMAttribute(17) = "TD-TI:TU"
#       Set rngBsMAttribute(17) = wksBsM.Cells(lngDatensatz, intZielspalte + 27)
#       strBsMAttribute(18) = "Testfälle"
#       Set rngBsMAttribute(18) = wksBsM.Cells(lngDatensatz, intZielspalte + 28)
#       strBsMAttribute(19) = "Vergleich TUs (TD-TF) - operativ"
#       Set rngBsMAttribute(19) = wksBsM.Cells(lngDatensatz, intZielspalte + 29)
#       strBsMAttribute(28) = "Erläuterungen zum Vergleich"
#       Set rngBsMAttribute(28) = wksBsM.Cells(lngDatensatz, intZielspalte + 30)
#       strBsMAttribute(26) = "Fehlende TUs bei TD-AAs"
#       Set rngBsMAttribute(26) = wksBsM.Cells(lngDatensatz, intZielspalte + 31)
#       strBsMAttribute(27) = "Fehlende TUs bei TFs"
#       Set rngBsMAttribute(27) = wksBsM.Cells(lngDatensatz, intZielspalte + 32)
#       strBsMAttribute(24) = "Zugeordnete I-Stufe"
#       Set rngBsMAttribute(24) = wksBsM.Cells(lngDatensatz, intZielspalte + 33)
#       strBsMAttribute(30) = "Projekt"
#       Set rngBsMAttribute(30) = wksBsM.Cells(lngDatensatz, intZielspalte + 34)
#       
#       'Ergänzung projektspezifischer Attribute

#       If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
        if not self.is_project_specific:
#           ReDim Preserve strBsMAttribute(LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1) + 1)
#           ReDim Preserve rngBsMAttribute(LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1))
#           strBsMAttribute(35) = "Temp11_Auswahlfeld"
            bsm_attributes = bsm_attributes.remove(OutputBSMAttribute.Temp11SelectionField)
#           Set rngBsMAttribute(35) = wksBsM.Cells(lngDatensatz, intZielspalte + 35)    ' => nachträglich an richtige Stelle verschieben?
#       End If
        bsm_output_data = {name : [] for name in bsm_attributes}
#       
        # Header format is implemented in Worksheet.append method
#       'Tabellenkopf anlegen
#       For i = LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1)
#           With rngBsMAttribute(i)
#               .Value = strBsMAttribute(i)
#               .Font.Bold = True
#               .Interior.Color = RGB(217, 217, 217)
#           End With
#       Next i
#       
#       'Bekannte Testumgebungen
#       ReDim strBekannteTUs(1 To 17)
#       intRelevantekTUs = 9
#       strBekannteTUs(1) = "BRS-HiL_Laborplatz_automatisiert"
#       strBekannteTUs(2) = "BRS-HiL_Basis-Funktion"
#       strBekannteTUs(3) = "BRS-HiL_Kunden-Funktion"
#       strBekannteTUs(4) = "BRS-HiL_Bremssystem"
#       strBekannteTUs(5) = "BRS-Fahrversuch_Kunden-Funktion"
#       strBekannteTUs(6) = "BRS-Fahrversuch_Basis-Funktion"
#       strBekannteTUs(7) = "Vernetzter-Fahrwerks-HiL_Kundenfunktion"
#       strBekannteTUs(8) = "BRS-HiL_Basisdienst_Halten"
#       strBekannteTUs(9) = "BRS-HiL_Basisdienst_Verzoegern"
#       'ab hier nicht mehr relevant
#       strBekannteTUs(10) = "BRS-SiL_Kunden-Funktion"
#       strBekannteTUs(11) = "Code-Review"
#       strBekannteTUs(12) = "Design-Review"
#       strBekannteTUs(13) = "Dokumenten-Review"
#       strBekannteTUs(14) = "Prozess-Review"
#       strBekannteTUs(15) = "Entscheidung_liegt_bei_Testinstanz"
#       strBekannteTUs(16) = "BRS-Fahrversuch_Applikation"
#       strBekannteTUs(17) = "BRS-Fahrversuch_Erprobung"
#       'Statuswerte intBekannteTUs:
#       '   TF \ VK                         kein VK     VK vorhanden
#       '   kein TF                         0           1
#       '   TF operativ                     10          11
#       '   TF nicht operativ               20          21
#       '   TF operativ und nicht operativ  30          31
#       
#       'Zähler für weitere Testumgebungen
#       intWeitereTUs = 0
        other_test_environment_top = 0
#       strWeitereTUsAusgabe = ""
        self.other_test_environment_output = ''
#       
#       'Tabelle füllen
#       'Relevante Testumgebungen für Abgleich zwischen TDs und TFs
#       ReDim strAbgleichTUs(1 To intRelevantekTUs)
#       For i = 1 To intRelevantekTUs
#           strAbgleichTUs(i) = strBekannteTUs(i)
#       Next i
        self.relevant_test_environments = KNOWN_TEST_ENVIRONMENTS[:RELEVANT_TOP]
#       
#       lngDatensatz = 0
        dataset_count = 0
#       'BsM-Daten ausgeben
#       For Each varErfassteBsMDatensatzItem In BsMDatenList
        bsm_dataset:BSMData|BSMSuccessorData
        for bsm_dataset in self.bsm_datasets.values():
            row_output = {att:'' for att in bsm_attributes}
#           'Zähler für Datensatz/Zeile
#           lngDatensatz = lngDatensatz + 1
            dataset_count += 1
#           'Kalenderwoche der Datenauswertung
#           If WorksheetFunction.WeekNum(Date, 2) < 10 Then
#               rngBsMAttribute(34).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & "0" & WorksheetFunction.WeekNum(Date, 2))      
#           Else
#               rngBsMAttribute(34).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & WorksheetFunction.WeekNum(Date, 2))
#           End If
            row_output[OutputBSMAttribute.KWDataEvaluation] = self.date_signature
#           'Vorgänger ID
#           If blnAVWVorgaengerIDsVerwenden Then
            if self.use_predecessor_ids:
#               rngBsMAttribute(0).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWVorgaengerID
                row_output[OutputBSMAttribute.RedirectedFrom] = bsm_dataset.avw_predecessor_id
            #else:
            #   row_output[OutputBSMAttribute.RedirectedFrom] = ''
#           End If
#           'Ausgabe ID
#           rngBsMAttribute(1).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWID
            row_output[OutputBSMAttribute.ID] = bsm_dataset.avw_id
#           'Ausgabe Dokument-ID
#           rngBsMAttribute(2).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWDokumentID
            row_output[OutputBSMAttribute.DocumentID] = bsm_dataset.avw_dokument_id
#           'Ausgabe BsM-Relevanz
#           rngBsMAttribute(3).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.BSMRelevanz
            row_output[OutputBSMAttribute.BSMRelevance] = bsm_dataset.bsm_relevanz
#           'Ausgabe BSM-SaFuSi
#           rngBsMAttribute(4).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMSaFuSi
            row_output[OutputBSMAttribute.BSMSaFuSiAssesment] = bsm_dataset.avw_bsm_safusi
#           'Ausgabe BSM-ZZ
#           rngBsMAttribute(5).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMZZ
            row_output[OutputBSMAttribute.BSMZZAssesment] = bsm_dataset.avw_bsm_zz
#           'Ausgabe BSM-ED
#           rngBsMAttribute(6).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMED
            row_output[OutputBSMAttribute.BSMEDAssesment] = bsm_dataset.avw_bsm_ed
#           'Ausgabe BSM-FFF
#           rngBsMAttribute(7).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMFFF
            row_output[OutputBSMAttribute.BSMFFFAssesment] = bsm_dataset.avw_bsm_fff
#           'Ausgabe BSM-O
#           rngBsMAttribute(8).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMO
            row_output[OutputBSMAttribute.BSMOAssesment] = bsm_dataset.avw_bsm_o
#           'Ausgabe BSM-Se
#           rngBsMAttribute(9).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMSe
            row_output[OutputBSMAttribute.BSMSeAssesment] = bsm_dataset.avw_bsm_se
#           'Ausgabe ASIL
#           rngBsMAttribute(10).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWASIL
            row_output[OutputBSMAttribute.ASIL] = bsm_dataset.avw_asil
#           'Ausgabe Feature
#           rngBsMAttribute(11).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWFeature
            row_output[OutputBSMAttribute.Feature] = bsm_dataset.avw_feature
#           'Ausgabe Reifegrad
#           rngBsMAttribute(12).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWReifegrad
            row_output[OutputBSMAttribute.MaturityLevel] = bsm_dataset.avw_reifegrad
#           'Ausgabe Umsetzer
#           rngBsMAttribute(13).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWUmsetzer
            row_output[OutputBSMAttribute.Implementer] = bsm_dataset.avw_implementer
#           'Ausgabe Status
#           rngBsMAttribute(14).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWStatus
            row_output[OutputBSMAttribute.Status] = bsm_dataset.avw_status
#           'Ausgabe MV
#           rngBsMAttribute(20).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWMV
            row_output[OutputBSMAttribute.MV] = bsm_dataset.avw_mv
#           'Ausgabe Kategorie
#           rngBsMAttribute(21).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWKategorie
            row_output[OutputBSMAttribute.Category] = bsm_dataset.avw_kategorie
#           'Ausgabe Dokumentenname
#           rngBsMAttribute(22).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWDokumentName
            row_output[OutputBSMAttribute.Document] = bsm_dataset.avw_dokument_name
#           'Ausgabe #abgelehnt_nicht_testbar
#           rngBsMAttribute(23).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWAbgelehntNichtTestbar
            row_output[OutputBSMAttribute.RejectedNotTestable] = bsm_dataset.avw_abgelehnt_nicht_testbar
#           'Ausgabe Zugeordnete I-Stufe
#           rngBsMAttribute(24).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.IStufe
            row_output[OutputBSMAttribute.AssignedILevel] = bsm_dataset.i_stufe
#           'Ausgabe Cluster Testing
#           rngBsMAttribute(29).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.ClusterTesting
            row_output[OutputBSMAttribute.TestingCluster] = bsm_dataset.cluster_testing
#           'Ausgabe Projekt
#           rngBsMAttribute(30).Offset(lngDatensatz, 0).Value = strProjekt
            row_output[OutputBSMAttribute.Project] = self.project
#           'Ausgabe Anforderungsverantwortliche
#           rngBsMAttribute(33).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWAnforderungsverantwortliche
            row_output[OutputBSMAttribute.RequirementOwner] = bsm_dataset.avw_anforderungsverantwortliche
#       
#           'Rücksetzen der Variablen für TU-Abgleich
#           ReDim intAbgleichTUs(LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1))
            self.te_comparison_count = [0] * len(self.relevant_test_environments)
#           ReDim intAuswertungTUs(1 To 31)
            self.test_environments_evaluation = [0] * 31
            te_evaluations = TestEnvironmentEvaluations(self.te_comparison_count)
#           ReDim strAuswertungTUs(1 To 31)
            #test environment (EN) = Testumgebung (DE)
            #security application (EN) = Absicherungsauftrag (DE)
#           strAuswertungTUsFehlendeAAs = ""
            #initialized in te_evaluations
#           strAuswertungTUsFehlendeTFs = ""
            #initialized in te_evaluations
#           intAusgabeAuswertungTUs = 0
            #initialized in te_evaluations
#           strAusgabeAuswertungTUs = ""
            #initialized in te_evaluations
#           strAusgabeAuswertungTUsDetails = ""
            #initialized in te_evaluations
#
#           'Auswertung TF
            str_test_cases = self.auswertung_tf(test_cases=bsm_dataset.test_cases)     
#           
#           'Auswertung TD
#           strTDAA = ""
            strTDAA = ''
#           strTDTiTu = ""
            strTDTiTu = ''  
#           blnTIUnerlaubt = False
#           
#           'Trennung der Umsetzer für Abgleich der Testinstanzen
#           varUmsetzer = Split(varErfassteBsMDatensatzItem.AVWUmsetzer, ",", , vbBinaryCompare)
            varUmsetzer = bsm_dataset.avw_implementer.split(",") if bsm_dataset.avw_implementer else []
#           If varErfassteBsMDatensatzItem.AVWUmsetzer <> "" Then
            if bsm_dataset.avw_implementer:
#               ReDim blnUmsetzer(LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1))
#               For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
#                   blnUmsetzer(intUmsetzer) = False
#               Next intUmsetzer
                bln_umsetzer = [False] * len(varUmsetzer)   
#           End If
#           
#           If varErfassteBsMDatensatzItem.Verifikationskriterium.Count > 0 Then
            if len(bsm_dataset.verifications_criteria) > 0:
#               'Ausgabe TD-VK inkl. Sicherheitsprüfung für mehrere Verifikationskriterien
                first_verification_criterion = bsm_dataset.verifications_criteria[0]
#               If varErfassteBsMDatensatzItem.Verifikationskriterium.Count > 1 Then
                if len(bsm_dataset.verifications_criteria) > 1:
#                   rngBsMAttribute(15).Offset(lngDatensatz, 0).Value = "Achtung, mehrere Verifikationskriterien vorhanden! Ausgabe nur des ersten Items." & vbCrLf & _
#                   varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_ID
                    row_output[OutputBSMAttribute.TDVC] = (
                        'Achtung, mehrere Verifikationskriterien vorhanden! Ausgabe nur des ersten Items.\n'
                        f'{first_verification_criterion.vk_id}'
                    )
#               Else
                else:
#                   rngBsMAttribute(15).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_ID
                    row_output[OutputBSMAttribute.TDVC] = first_verification_criterion.vk_id
#               End If
#               'Ausgabe TD-VK Status
#               rngBsMAttribute(25).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_status
                row_output[OutputBSMAttribute.StatusTDVC] = first_verification_criterion.status
#               'Ausgabe TD-VK temp1_Text
#               rngBsMAttribute(31).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_temp1Text
                row_output[OutputBSMAttribute.TDVCTemp1Text] = first_verification_criterion.temp1_text
#               'Ausgabe TD-VK Effort Estimation - Aufwandsschätzung auf Basis der Vorkommen von "Use-Case", "Step", "Aktion"
#               dblTDVKAnzahlUseCases = 1
                dblTDVKAnzahlUseCases = 1
#               strTDVKAktion = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_Aktion
                strTDVKAktion = first_verification_criterion.aktion
#               If strTDVKAktion <> "" Then
                if strTDVKAktion != '':
#                   strTDVKAktion = Replace(UCase(strTDVKAktion), "USE CASE", "USE-CASE")
#                   strTDVKAktion = Replace(UCase(strTDVKAktion), "USECASE", "USE-CASE")
                    strTDVKAktion = re.sub(r'USE( )?CASE', 'USE-CASE',strTDVKAktion.upper())
#                   dblTDVKAnzahlUseCases = (Len(strTDVKAktion) - Len(Replace(UCase(strTDVKAktion), "USE-CASE", ""))) / Len("Use-Case")
                    dblTDVKAnzahlUseCases = (len(strTDVKAktion) - len(re.sub("USE-CASE","", strTDVKAktion)))// len('Use-Case')
#                   'Anzahl 1 bei Befüllung ohne Vorkommen der Schlagwörter
#                   If dblTDVKAnzahlUseCases = 0 Then dblTDVKAnzahlUseCases = 1
                    if dblTDVKAnzahlUseCases == 0:
                        dblTDVKAnzahlUseCases = 1
#               End If
#               rngBsMAttribute(32).Offset(lngDatensatz, 0).Value = dblTDVKAnzahlUseCases
                row_output[OutputBSMAttribute.TDVCEffortEstimation] = dblTDVKAnzahlUseCases
#              
#               'Auswertung TD-AA
#               If varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).Absicherungsauftraege.Count > 0 Then
                if len(first_verification_criterion.absicherungsauftraege) > 0:
                    security_orders = list(first_verification_criterion.absicherungsauftraege.values())
                    #------1)
                    strTDAA = CRLF.join([so.abs_id for so in security_orders])
                    #------2)
                    strTDTiTu = CRLF.join([f'{so.test_instance}: {so.test_environment_type}' for so in security_orders])
#                   For Each varErfassteTDAAItem In varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).Absicherungsauftraege
                    for security_order in security_orders:
                        #----------------------------------------------- implemented above 1)
#                       'Absicherungsaufträge zusammenführen
#                       If strTDAA = "" Then
#                           strTDAA = varErfassteTDAAItem.abs_ID
                            #implemented in if sentence above this for loop
#                       Else
#                           strTDAA = strTDAA & vbCrLf & varErfassteTDAAItem.abs_ID
#                       End If
                        #----------------------------------------------- implemented above 2)
#                       'Ti-Tu-Kombinationen zusammenführen
#                       If strTDTiTu = "" Then
#                           strTDTiTu = varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
                            #implemented in if sentence above this for loop
#                       Else
#                           strTDTiTu = strTDTiTu & vbCrLf & varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                       End If
                        #-----------------------------------------------
#                       
#                       'Auswertung ob relevante Testinstanzen abgedeckt sind
#                       If varErfassteTDAAItem.testinstanz <> "eigene Organisationseinheit" And varErfassteTDAAItem.testinstanz <> "Dauerlauf Gesamtfahrzeug" And varErfassteTDAAItem.testinstanz <> "Gesamtverbundintegration" And varErfassteTDAAItem.testinstanz <> "HMS" And varErfassteTDAAItem.testinstanz <> "VZM" Then
                        if security_order.test_instance not in ['Dauerlauf Gesamtfahrzeug', 'Gesamtverbundintegration', 'HMS', 'VZM']:
#                           For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
                            for index, implementer in enumerate(varUmsetzer):
#                               If Trim(varErfassteTDAAItem.testinstanz) = Trim(varUmsetzer(intUmsetzer)) Then
                                if security_order.test_instance.strip() == implementer.strip():
#                                   blnUmsetzer(intUmsetzer) = True
                                    bln_umsetzer[index] = True
#                                   Exit For
                                    break
#                               End If
#                           Next intUmsetzer                           
#                       End If
#                       
#       '                'Auswertung Testinstanz Erlaubt/Unerlaubt
#       '                If varErfassteTDAAItem.testinstanz <> "eigene Organisationseinheit" And varErfassteTDAAItem.testinstanz <> "Dauerlauf Gesamtfahrzeug" And varErfassteTDAAItem.testinstanz <> "Gesamtverbundintegration" And varErfassteTDAAItem.testinstanz <> "HMS" And varErfassteTDAAItem.testinstanz <> "VZM" Then
#                           For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
#       '                    If AbgleichUmsetzerTI(varErfassteTDAAItem.testinstanz, varErfassteBsMDatensatzItem.AVWUmsetzer) = False Then
#       '                        blnTIUnerlaubt = True
#       '                    End If
#       '                End If
#
#       'Abgleich der vorhandenen relevanten Testumgebungen                       
                        self.abgleich_TUs(security_order)
#                   Next varErfassteTDAAItem
#               
#                   'Auswertung des Abgleichs der relevanten Testumgebungen
                    #Evaluation of the comparison of relevant test environments
#                   Call AuswertungTUAbgleich(intAbgleichTUs, strAbgleichTUs, intAuswertungTUs, strAuswertungTUs)
                    #te_evaluations = TestEnvironmentEvaluations(self.te_comparison_count)
                    te_evaluations.summarize()
#                   
#                   'Erzeugung der Ausgabe für TU-Vergleich
#                   Call AusgabeTUAbgleich(intAuswertungTUs, strAuswertungTUs, strAuswertungTUsFehlendeAAs, strAuswertungTUsFehlendeTFs, intAusgabeAuswertungTUs, strAusgabeAuswertungTUs, strAusgabeAuswertungTUsDetails)
                    te_evaluations.output_comparison()
#               Else
                else:
#                   'Keine Absicherungsaufträge vorhanden
#                   strAusgabeAuswertungTUs = "Kein Absicherungsauftrag vorhanden"
#                   intAusgabeAuswertungTUs = 3
                    te_evaluations.empty('Kein Absicherungsauftrag vorhanden')
#               End If
#           Else
            else:
#               'Kein Verifikationskriterium vorhanden
#               strAusgabeAuswertungTUs = "Kein Verifikationskriterium vorhanden"
#               intAusgabeAuswertungTUs = 3
                te_evaluations.empty('Kein Verifikationskriterium vorhanden')
#           End If
#           
#           'Ausgabe Testfälle
#           rngBsMAttribute(18).Offset(lngDatensatz, 0).Value = strTestfaelle
            row_output[OutputBSMAttribute.TestCases] = str_test_cases
#           'Ausgabe TD-AA
#           rngBsMAttribute(16).Offset(lngDatensatz, 0).Value = strTDAA
            row_output[OutputBSMAttribute.TDSafeguards] = strTDAA
#           'Ausgabe TD-TI:TU
#           rngBsMAttribute(17).Offset(lngDatensatz, 0).Value = strTDTiTu
            row_output[OutputBSMAttribute.TDTITE] = strTDTiTu
#           'Alle Umsetzer in den Testinstanzen abgedeckt?
#           blnTIUnerlaubt = False
            blnTIUnerlaubt = any(not implementer for implementer in bln_umsetzer)
#           For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
#               If blnUmsetzer(intUmsetzer) = False Then
#                   blnTIUnerlaubt = True
#               End If
#           Next intUmsetzer
            row = len(bsm_output_data[OutputBSMAttribute.TDTITE]) + 2 # dataframe row 0 is in excel row 2
#           If blnTIUnerlaubt Then
            if blnTIUnerlaubt:
#               rngBsMAttribute(17).Offset(lngDatensatz, 0).Interior.Color = RGB(255, 255, 102)
                bg_colours[OutputBSMAttribute.TDTITE.value][row] = BG_YELLOW
#           End If
#           'Ausgabe Vergleich TUs
#           With rngBsMAttribute(19).Offset(lngDatensatz, 0)
#               .Value = strAusgabeAuswertungTUs
            row_output[OutputBSMAttribute.OperationalComparisonTEsTDTC] = te_evaluations.output
#               If intAusgabeAuswertungTUs = 1 Then
            if te_evaluations.int_output == 1:
#                   'Grün
#                   .Interior.Color = RGB(51, 204, 51)
                bg_colours[OutputBSMAttribute.OperationalComparisonTEsTDTC.value][row] = BG_GREEN
                bg_colours[OutputBSMAttribute.ComparisonExplanations.value][row] = BG_GREEN
#               ElseIf intAusgabeAuswertungTUs = 2 Then
            elif te_evaluations.int_output == 2:
#                   'Gelb
#                   .Interior.Color = RGB(255, 255, 102)
                bg_colours[OutputBSMAttribute.OperationalComparisonTEsTDTC.value][row] = BG_YELLOW
                bg_colours[OutputBSMAttribute.ComparisonExplanations.value][row] = BG_YELLOW
#               ElseIf intAusgabeAuswertungTUs = 3 Then
            elif te_evaluations.int_output == 3:
#                   'Rot
#                   .Interior.Color = RGB(255, 51, 0)
                bg_colours[OutputBSMAttribute.OperationalComparisonTEsTDTC.value][row] = BG_RED
                bg_colours[OutputBSMAttribute.ComparisonExplanations.value][row] = BG_RED
#               End If
            ###-------------------------------------------------
#           End With
#           With rngBsMAttribute(28).Offset(lngDatensatz, 0)
#               .Value = strAusgabeAuswertungTUsDetails
            row_output[OutputBSMAttribute.ComparisonExplanations] = te_evaluations.output_details
            # Background colours implemented above
#                   'Grün
#                   .Interior.Color = RGB(51, 204, 51)
#                   'Gelb
#                   .Interior.Color = RGB(255, 255, 102)
#               ElseIf intAusgabeAuswertungTUs = 3 Then
#                   'Rot
#                   .Interior.Color = RGB(255, 51, 0)
#               End If
            ###-------------------------------------------------
#           End With
#           'Ausgabe Erläuterungen zum Vergleich
#           'Ausgabe fehlende TUs bei TD-AAs
#           rngBsMAttribute(26).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeAAs
            row_output[OutputBSMAttribute.MissingTEInTDSafeguards] = te_evaluations.missing_safe_guards
#           'Ausgabe fehlende TUs bei TFs
#           rngBsMAttribute(27).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeTFs
            row_output[OutputBSMAttribute.MissingTEInTCs] = te_evaluations.missing_test_cases
#           
#           'Projektspezifische Ausgabe - MEB21
#           If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
            if self.project in ['MEB21', 'MQB48W']:
#               rngBsMAttribute(35).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWTemp11_Auswahlfeld
                row_output[OutputBSMAttribute.Temp11SelectionField] = bsm_dataset.avw_temp11_auswahlfeld
#           End If
            for field, value in row_output.items():
                bsm_output_data[field].append(value)
#           
#       Next varErfassteBsMDatensatzItem
#       
#       'Weitere TUs zusammenfassen
#       If intWeitereTUs > 0 Then
        if len(self.other_test_environment_output) > 0:
            self.other_test_environment_output = CRLF.join(self.other_test_environment)
#           For i = LBound(strWeitereTUs, 1) To UBound(strWeitereTUs, 1)
            #for other_te_i in range(other_test_environment_top):
#               If strWeitereTUsAusgabe = "" Then
                #if self.other_test_environment_output == '':
#                   strWeitereTUsAusgabe = strWeitereTUs(i)
                    #self.other_test_environment_output = strWeitereTUs[other_te_i]
#               Else
                #else:
#                   strWeitereTUsAusgabe = strWeitereTUsAusgabe & vbCrLf & strWeitereTUs(i)
                #self.other_test_environment_output += '\n' + strWeitereTUs[other_te_i]
#               End If
#           Next i
#       End If
#       
#       'Spaltenbreite anpassen
        # cells width formating is implemented in Workbook.append method
#       With wksBsM.Cells
#           .Columns.AutoFit
#           .Rows.AutoFit
#       End With
#       
#       'Projektspezifische Sortierung - MEB21
        # Implemented by declaring the column names sorted as wanted
        ###Done at bsm_output_data definition
#       If strProjekt = "MEB21" Then
#           wksBsM.Columns(35).Cut
#           wksBsM.Columns(18).Insert shift:=xlToRight
#       End If
#       
        ###TODO activate Filtering
#       'Filterung aktivieren
#       wksBsM.Rows(1).AutoFilter
#       
#       'Dateinamen ausgeben und verstecken
        ###TODO
#       lngDatensatz = 1
#       wksBsM.Rows(lngDatensatz).EntireRow.Insert shift:=xlDown
#       wksBsM.Cells(lngDatensatz, 1) = "Anforderungen: " & strDateinamen(1) & vbCrLf & "Verifikationskriterien: " & strDateinamen(2) & vbCrLf & _
#                                       "Absicherungsaufträge: " & strDateinamen(3) & vbCrLf & "Testfälle: " & strDateinamen(3) & vbCrLf & _
#                                       "FRU-Timing: " & strDateinamen(5)
#       wksBsM.Rows(lngDatensatz).EntireRow.Hidden = True
        #self.bsm_output_data = bsm_output_data
        #self.output_workbook.append_worksheet(
        #    data_frame=pd.DataFrame(
        #        bsm_output_data,
        #        dtype=str
        #    ),
        #    name=worksheet_name,
        #)
        return bsm_output_data, worksheet_name, bg_colours
#   End Sub

    def output_worksheets(self):
        yield self.output_status()
        yield self.output_status_TD()
#   
#   Private Sub AusgabeTDStatus(ByVal wbBsM As Workbook, ByRef wksTD As Worksheet, ByRef strTDAttribute() As String, ByRef rngTDAttribute() As Range, ByRef strDateinamen() As String, ByVal strProjekt As String)
    def output_status_TD(self) -> tuple[dict[str,list[str]],str]:
#       Dim lngDatensatz As Long                        'Long-Variable für aktuell zu schreibenden Datensatz
#       Dim Verifikationskriterium As Verifikationskriterium    'Verifikationskriterium
#       Dim varErfassteTDAAItem As Variant              'Variant für Item aus den jeweiligen Absicherungsaufträgen
#       Dim strTDAA As String                           'String für Sammlung der Absicherungsaufträge
#       Dim strTDTiTu As String                         'String für Sammlung der Ti:Tu-Kombinationen
#       Dim varErfassteTFItem As Variant                'Variant für Item aus den jeweiligen Testfällen
#       Dim strTestfaelle As String                     'String für Sammlung der Testfälle
#       Dim strAbgleichTUs() As String                  'String-Array für die Namen der abzugleichenden Testumgebungen bei TDs und TFs
#       Dim intAbgleichTUs() As Integer                 'Integer-Array für Erfassung der Testumgebungen bei TDs und TFs
#       Dim i As Long                                   'Laufvariable
#       Dim intAuswertungTUs() As Integer               'Integer-Array für die Ergebnisse des Tu-Abgleichs
#       Dim strAuswertungTUs() As String                'String-Array für die Ergebnisse des Tu-Abgleichs
#       Dim strAusgabeAuswertungTUs As String           'String für Ausgabe des Tu-Abgleichs
#       Dim intAusgabeAuswertungTUs As Integer          'Integer für Ausgabe des Tu-Abgleichs
#       Dim strAuswertungTUsFehlendeAAs As String       'String für Ausgabe der fehlenden TUs bei TD-AAs
#       Dim strAuswertungTUsFehlendeTFs As String       'String für Ausgabe der fehlenden TUs bei TFs
#       Dim strAusgabeAuswertungTUsDetails As String    'String für Ausgabe des Tu-Abgleichs mit Details
#       Dim intWeitereTUs As Integer                    'Zählvariable für weitere TUs
#       Dim strWeitereTUs() As String                   'String-Array für weitere TUs
        self.other_test_environment = []
#       Dim strBekannteTUs() As String                  'String-Array für alle bekannten TUs
#       Dim blnTFTUZugeordnet As Boolean                'Flag für zugeordnete TU
#       Dim blnAATUZugeordnet As Boolean                'Flag für zugeordnete TU
#       Dim intWeitereTUsZaehler As Integer             'Laufvariable für weitere TUs
#       Dim intRelevantekTUs As Integer                 'Integer für Anzahl der relevanten TUs
#       Dim dblTDVKAnzahlUseCases As Double             'Anzahl der Vorkommen der Use-Case-Begriffe
#       Dim strTDVKAktion As String                     'String zur Bearbeitung der TDVK-Aktion
        bg_colours:FillDict = FillDict()
#       
#       'Tabelle erzeugen
#       'Neues Worksheet erzeugen
#       Set wksTD = wbBsM.Sheets.Add(after:=wbBsM.Worksheets(wbBsM.Worksheets.Count))
#       wksTD.Name = "TD_Status_" & "Today" & "_" & Replace(Time, ":", "")
        worksheet_name = f'TD_Status_{self.date_suffix}'
#       'Arbeitsblatt TD_Status
#       '#1: TD-VK, #2: Status TD-VK, #3: TD-AA, #4: TD-TI:TU, #5: Testfälle, #6: Vergleich TUs (TD-TF) - operativ, #7: Erläuterungen zum Vergleich,
#       '#8: Fehlende TUs bei TD-AAs, #9: Fehlende TUs bei TFs, #10: Anforderungs-IDs, #11: Zugeordnete I-Stufe, #12: Umsetzer, #13: BsM-Relevanz,
#       '#14: ASIL, #15: Feature, #16: Reifegrad, #17: MV, #18: LAH-ID, #19: Dokumente (LAH), #20: Cluster Testing, #21: Projekt, #22: TD-VK temp1_Text,
#       '#23: TD-VK Effort Estimation, #24: Anforderungsverantwortliche, #25: KW
#       'Projektspezifisch - MEB21
#       '#26: Temp11_Auswahlfeld
#       ReDim strTDAttribute(1 To 25)
#       ReDim rngTDAttribute(LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1))
#       'Name und Position der Tabellenattribute
#       strTDAttribute(25) = "KW Datenauswertung"
#       Set rngTDAttribute(25) = wksTD.Cells(1, 1)
#       strTDAttribute(1) = "TD-VK"
#       Set rngTDAttribute(1) = wksTD.Cells(1, 2)
#       strTDAttribute(2) = "Status TD-VK"
#       Set rngTDAttribute(2) = wksTD.Cells(1, 3)
#       strTDAttribute(22) = "TD-VK temp1_Text"
#       Set rngTDAttribute(22) = wksTD.Cells(1, 4)
#       strTDAttribute(23) = "TD-VK Effort Estimation"
#       Set rngTDAttribute(23) = wksTD.Cells(1, 5)
#       strTDAttribute(3) = "TD-AA"
#       Set rngTDAttribute(3) = wksTD.Cells(1, 6)
#       strTDAttribute(4) = "TD-TI:TU"
#       Set rngTDAttribute(4) = wksTD.Cells(1, 7)
#       strTDAttribute(5) = "Testfälle"
#       Set rngTDAttribute(5) = wksTD.Cells(1, 8)
#       strTDAttribute(6) = "Vergleich TUs (TD-TF) - operativ"
#       Set rngTDAttribute(6) = wksTD.Cells(1, 9)
#       strTDAttribute(7) = "Erläuterungen zum Vergleich"
#       Set rngTDAttribute(7) = wksTD.Cells(1, 10)
#       strTDAttribute(8) = "Fehlende TUs bei TD-AAs"
#       Set rngTDAttribute(8) = wksTD.Cells(1, 11)
#       strTDAttribute(9) = "Fehlende TUs bei TFs"
#       Set rngTDAttribute(9) = wksTD.Cells(1, 12)
#       strTDAttribute(10) = "Anforderungs-IDs"
#       Set rngTDAttribute(10) = wksTD.Cells(1, 13)
#       strTDAttribute(20) = "Cluster Testing"
#       Set rngTDAttribute(20) = wksTD.Cells(1, 14)
#       strTDAttribute(14) = "ASIL (LAH)"
#       Set rngTDAttribute(14) = wksTD.Cells(1, 15)
#       strTDAttribute(13) = "BsM-Relevanz (LAH)"
#       Set rngTDAttribute(13) = wksTD.Cells(1, 16)
#       strTDAttribute(15) = "Feature (LAH)"
#       Set rngTDAttribute(15) = wksTD.Cells(1, 17)
#       strTDAttribute(16) = "Reifegrad (LAH)"
#       Set rngTDAttribute(16) = wksTD.Cells(1, 18)
#       strTDAttribute(12) = "Umsetzer (LAH)"
#       Set rngTDAttribute(12) = wksTD.Cells(1, 19)
#       strTDAttribute(17) = "MV (LAH)"
#       Set rngTDAttribute(17) = wksTD.Cells(1, 20)
#       strTDAttribute(24) = "Anforderungsverantwortliche (LAH)"
#       Set rngTDAttribute(24) = wksTD.Cells(1, 21)
#       strTDAttribute(18) = "LAH-ID"
#       Set rngTDAttribute(18) = wksTD.Cells(1, 22)
#       strTDAttribute(19) = "Dokumente (LAH)"
#       Set rngTDAttribute(19) = wksTD.Cells(1, 23)
#       strTDAttribute(11) = "Zugeordnete I-Stufe"
#       Set rngTDAttribute(11) = wksTD.Cells(1, 24)
#       strTDAttribute(21) = "Projekt"
#       Set rngTDAttribute(21) = wksTD.Cells(1, 25)
        #TD-Attribute collection
        td_attributes = [attribute.value for attribute in TDAttribute]
#       
#       'Ergänzung projektspezifische Attribute
#       If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
        if not self.is_project_specific:
#           ReDim Preserve strTDAttribute(LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1) + 1)
#           ReDim Preserve rngTDAttribute(LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1))
#           strTDAttribute(26) = "Temp11_Auswahlfeld (LAH)"
            #it should be at 13 column
            td_attributes = td_attributes.remove(TDProjectAttribute.Temp11SelectionField)
            ####td_attributes.append(TDProjectAttribute.Temp11SelectionField)
#           Set rngTDAttribute(26) = wksTD.Cells(1, 26)
#       End If
#       
        # Implemented in Workbook.append method
#       'Tabellenkopf anlegen
#       For i = LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1)
#           With rngTDAttribute(i)
#               .Value = strTDAttribute(i)
#               .Font.Bold = True
#               .Interior.Color = RGB(217, 217, 217)
#           End With
#       Next i
#       
#       'Bekannte Testumgebungen
#       ReDim strBekannteTUs(1 To 17)
#       intRelevantekTUs = 9
#       strBekannteTUs(1) = "BRS-HiL_Laborplatz_automatisiert"
#       strBekannteTUs(2) = "BRS-HiL_Basis-Funktion"
#       strBekannteTUs(3) = "BRS-HiL_Kunden-Funktion"
#       strBekannteTUs(4) = "BRS-HiL_Bremssystem"
#       strBekannteTUs(5) = "BRS-Fahrversuch_Kunden-Funktion"
#       strBekannteTUs(6) = "BRS-Fahrversuch_Basis-Funktion"
#       strBekannteTUs(7) = "Vernetzter-Fahrwerks-HiL_Kundenfunktion"
#       strBekannteTUs(8) = "BRS-HiL_Basisdienst_Halten"
#       strBekannteTUs(9) = "BRS-HiL_Basisdienst_Verzoegern"
#       'ab hier nicht mehr relevant
#       strBekannteTUs(10) = "BRS-SiL_Kunden-Funktion"
#       strBekannteTUs(11) = "Code-Review"
#       strBekannteTUs(12) = "Design-Review"
#       strBekannteTUs(13) = "Dokumenten-Review"
#       strBekannteTUs(14) = "Prozess-Review"
#       strBekannteTUs(15) = "Entscheidung_liegt_bei_Testinstanz"
#       strBekannteTUs(16) = "BRS-Fahrversuch_Applikation"
#       strBekannteTUs(17) = "BRS-Fahrversuch_Erprobung"
#       'Statuswerte intBekannteTUs:
#       '   TF \ VK                         kein VK     VK vorhanden
#       '   kein TF                         0           1
#       '   TF operativ                     10          11
#       '   TF nicht operativ               20          21
#       '   TF operativ und nicht operativ  30          31
#       
#       'Zähler für weitere Testumgebungen
#       intWeitereTUs = 0
#       
#       'Tabelle füllen
#       'Relevante Testumgebungen für Abgleich zwischen TDs und TFs
#       ReDim strAbgleichTUs(1 To intRelevantekTUs)
#       For i = 1 To intRelevantekTUs
#           strAbgleichTUs(i) = strBekannteTUs(i)
#       Next i
#       
        self.relevant_test_environments = KNOWN_TEST_ENVIRONMENTS[:RELEVANT_TOP]
        td_output_data = {name: [] for name in td_attributes}

#       'TD-Daten ausgeben
#       lngDatensatz = 0
        dataset_count = 0

#       For Each Verifikationskriterium In verifikationKritList
        verification_criterion:Verificationskriterium
        for verification_criterion in self.verification_criteria.values():
            row_data = {name:'' for name in td_attributes}
#           If Verifikationskriterium.AnforderungVorhanden = True Then
            if verification_criterion.requirement_present:
#               'Zähler für Datensatz/Zeile
#               lngDatensatz = lngDatensatz + 1
                dataset_count += 1
#               'Kalenderwoche der Datenauswertung
#               If WorksheetFunction.WeekNum(Date, 2) < 10 Then
#                   rngTDAttribute(25).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & "0" & WorksheetFunction.WeekNum(Date, 2))
#               Else
#                   rngTDAttribute(25).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & WorksheetFunction.WeekNum(Date, 2))
#               End If
                row_data[TDAttribute.KWDataAnalysis] = self.date_signature
#               'Ausgabe VK-ID
#               rngTDAttribute(1).Offset(lngDatensatz, 0).Value = Verifikationskriterium.VK_ID
                row_data[TDAttribute.TDVC] = verification_criterion.vk_id
#               'Ausgabe VK-Status
#               rngTDAttribute(2).Offset(lngDatensatz, 0).Value = Verifikationskriterium.VK_status
                row_data[TDAttribute.StatusTDVC] = verification_criterion.status
#               'Ausgabe VK temp1_Text
#               rngTDAttribute(22).Offset(lngDatensatz, 0).Value = Verifikationskriterium.VK_temp1Text
                row_data[TDAttribute.TDVCTemp1Text] = verification_criterion.temp1_text
#               'Ausgabe Anforderungs-IDs
#               rngTDAttribute(10).Offset(lngDatensatz, 0).Value = AusgabeSammlungLF(Verifikationskriterium.anf_ids)
                row_data[TDAttribute.RequirementIDs] = CRLF.join(verification_criterion.requirement_ids)
#               'Ausgabe Zugeordnete I-Stufe
#               rngTDAttribute(11).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_IStufen)
                row_data[TDAttribute.AsignedILevel] = unic_join(CRLF, verification_criterion.requirement_i_level)
                row = len(td_output_data[TDAttribute.AsignedILevel]) + 2 # dataframe row 0 is in excel row 2
#               If AuswertungUnterschiedlicheIStufen(Verifikationskriterium.anf_IStufen) = True Then
                if self.different_i_levels(verification_criterion.requirement_i_level):
#                   rngTDAttribute(11).Offset(lngDatensatz, 0).Interior.Color = RGB(255, 255, 102)
                    bg_colours[TDAttribute.AsignedILevel.value][row] = BG_YELLOW
#               End If
#               'Ausgabe Umsetzer
#               rngTDAttribute(12).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Umsetzer)
                row_data[TDAttribute.LAH_Implementer] = unic_join(CRLF, verification_criterion.requirement_implementer)
#               'Ausgabe BsM-Relevanz
#               rngTDAttribute(13).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_BsMRelevanz)
                row_data[TDAttribute.LAH_BsMRelevance] = unic_join(CRLF, verification_criterion.requirement_bsm_relevance)
#               'Ausgabe ASIL
#               rngTDAttribute(14).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_ASIL)
                row_data[TDAttribute.LAH_ASIL] = unic_join(CRLF, verification_criterion.requirement_asil)
#               'Ausgabe Feature
#               rngTDAttribute(15).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Feature)
                row_data[TDAttribute.LAH_Feature] = unic_join(CRLF, verification_criterion.requirement_feature)
#               'Ausgabe Reifegrad
#               rngTDAttribute(16).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Reifegrad)
                row_data[TDAttribute.LAH_MaturityLevel] = unic_join(CRLF, verification_criterion.requirement_maturity_level)
#               'Ausgabe Modulverantwortlicher
#               rngTDAttribute(17).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_MV)
                row_data[TDAttribute.LAH_MV] = unic_join(CRLF, verification_criterion.requirement_mv)
#               'Ausgabe LAH-IDs
#               rngTDAttribute(18).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_LAHID)
                row_data[TDAttribute.LAH_ID] = unic_join(CRLF, verification_criterion.requirement_lah_id)
#               'Ausgabe LAH-Namen
#               rngTDAttribute(19).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_LAHNamen)
                row_data[TDAttribute.LAH_Document] = unic_join(CRLF, verification_criterion.requirement_lah_name)
#               'Ausgabe Cluster Testing
#               rngTDAttribute(20).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_ClusterTesting)
                row_data[TDAttribute.TestingCluster] = unic_join(CRLF, verification_criterion.requirement_cluster_testing)
#               'Ausgabe Projekt
#               rngTDAttribute(21).Offset(lngDatensatz, 0).Value = strProjekt
                row_data[TDAttribute.Project] = self.project
#               'Ausgabe Anforderungsverantwortliche
#               rngTDAttribute(24).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Anforderungsverantwortliche)
                row_data[TDAttribute.LAH_RequirementOwner] = unic_join(CRLF, verification_criterion.requirement_owner)
#               
#               'Ausgabe Aufwandsschätzung auf Basis der Vorkommen von "Use-Case", "Step", "Aktion"
#               dblTDVKAnzahlUseCases = 1
                use_cases_count = 1
#               strTDVKAktion = Verifikationskriterium.VK_Aktion
                tdvc_action:str = str(verification_criterion.aktion)
#               If strTDVKAktion <> "" Then
                if tdvc_action != '':
#                   strTDVKAktion = Replace(UCase(strTDVKAktion), "USE CASE", "USE-CASE")
#                   strTDVKAktion = Replace(UCase(strTDVKAktion), "USECASE", "USE-CASE")
                    tdvc_action = re.sub(r'USE( )?CASE', 'USE-CASE', tdvc_action.upper())
#                   dblTDVKAnzahlUseCases = (Len(strTDVKAktion) - Len(Replace(UCase(strTDVKAktion), "USE-CASE", ""))) / Len("Use-Case")
                    use_cases_count = (len(tdvc_action)-len(re.sub(r'USE-CASE','',tdvc_action))) // len('USE-CASE')
#                   'Anzahl 1 bei Befüllung ohne Vorkommen der Schlagwörter
#                   If dblTDVKAnzahlUseCases = 0 Then dblTDVKAnzahlUseCases = 1
                    if use_cases_count == 0:
                        use_cases_count = 1
#               End If
#               rngTDAttribute(23).Offset(lngDatensatz, 0).Value = dblTDVKAnzahlUseCases
                row_data[TDAttribute.TDVCEffortEstimation] = str(use_cases_count)
#               
#               'Rücksetzen der Variablen für TU-Abgleich
#               ReDim intAbgleichTUs(LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1))               
                self.te_comparison_count = [0] * len(self.relevant_test_environments)
#               ReDim intAuswertungTUs(1 To 31)
                self.test_environments_evaluation = [0] * 31
                te_evaluations = TestEnvironmentEvaluations(self.te_comparison_count)
#               ReDim strAuswertungTUs(1 To 31)
#               strAuswertungTUsFehlendeAAs = ""
                #initialized in te_evaluations
#               strAuswertungTUsFehlendeTFs = ""
                #initialized in te_evaluations
#               intAusgabeAuswertungTUs = 0
                #initialized in te_evaluations
#               strAusgabeAuswertungTUs = ""
                #initialized in te_evaluations
#               strAusgabeAuswertungTUsDetails = ""
                #initialized in te_evaluations
#           
#               'Auswertung TF
                str_test_cases = self.auswertung_tf(test_cases=list(verification_criterion.test_cases.values()))
#               
#               'Auswertung TD
#               strTDAA = ""
                strTDAA = ''
#               strTDTiTu = ""
                strTDTiTu = ''
#                   
#               'Auswertung TD-AA
#               If Verifikationskriterium.Absicherungsauftraege.Count > 0 Then
                if len(verification_criterion.absicherungsauftraege) > 0:
                    security_orders = list(verification_criterion.absicherungsauftraege.values())
                    #------1)
                    strTDAA = CRLF.join([so.abs_id for so in security_orders])
                    #------2)
                    strTDTiTu = CRLF.join([f'{so.test_instance}: {so.test_environment_type}' for so in security_orders])
#                   For Each varErfassteTDAAItem In Verifikationskriterium.Absicherungsauftraege
                    for security_order in security_orders:
                        #----------------------------------------------- implemented above 1)
#                       'Absicherungsaufträge zusammenführen
#                       If strTDAA = "" Then
#                           strTDAA = varErfassteTDAAItem.abs_ID
#                       Else
#                           strTDAA = strTDAA & vbCrLf & varErfassteTDAAItem.abs_ID
#                       End If
                        #----------------------------------------------- implemented above 2)
#                       'Ti-Tu-Kombinationen zusammenführen
#                       If strTDTiTu = "" Then
#                           strTDTiTu = varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                       Else
#                           strTDTiTu = strTDTiTu & vbCrLf & varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                       End If
                        #-----------------------------------------------
#
#       'Abgleich der vorhandenen relevanten Testumgebungen                         
                        self.abgleich_TUs(security_order)
#                   Next varErfassteTDAAItem
#                   
#                   'Auswertung des Abgleichs der relevanten Testumgebungen
#                   Call AuswertungTUAbgleich(intAbgleichTUs, strAbgleichTUs, intAuswertungTUs, strAuswertungTUs)
                    #te_evaluations = TestEnvironmentEvaluations(self.te_comparison_count)
                    te_evaluations.summarize()
#                   
#                   'Erzeugung der Ausgabe für TU-Vergleich
#                   Call AusgabeTUAbgleich(intAuswertungTUs, strAuswertungTUs, strAuswertungTUsFehlendeAAs, strAuswertungTUsFehlendeTFs, intAusgabeAuswertungTUs, strAusgabeAuswertungTUs, strAusgabeAuswertungTUsDetails)
                    te_evaluations.output_comparison()
#               Else
                else:
#                   'Keine Absicherungsaufträge vorhanden
#                   strAusgabeAuswertungTUs = "Kein Absicherungsauftrag vorhanden"
#                   intAusgabeAuswertungTUs = 3
                    te_evaluations.empty('Kein Absicherungsauftrag vorhanden')
#               End If
#                    
#               'Ausgabe TD-AA
#               rngTDAttribute(3).Offset(lngDatensatz, 0).Value = strTDAA
                row_data[TDAttribute.TDSafeGuards] = strTDAA
#               'Ausgabe Testfälle
#               rngTDAttribute(5).Offset(lngDatensatz, 0).Value = strTestfaelle
                row_data[TDAttribute.TestCases] = str_test_cases
#               'Ausgabe TD-TI:TU
#               rngTDAttribute(4).Offset(lngDatensatz, 0).Value = strTDTiTu
                row_data[TDAttribute.TDTITE] = strTDTiTu
#               'Ausgabe Vergleich TUs
#               With rngTDAttribute(6).Offset(lngDatensatz, 0)
#                   .Value = strAusgabeAuswertungTUs
                row_data[TDAttribute.OperativeTEComparisonTDTC] = te_evaluations.output
                ###TODO format cell background color
#                   If intAusgabeAuswertungTUs = 1 Then
                if te_evaluations.int_output == 1:
#                       'Grün
#                       .Interior.Color = RGB(51, 204, 51)
                    bg_colours[TDAttribute.OperativeTEComparisonTDTC.value][row] = BG_GREEN
                    bg_colours[TDAttribute.ComparisonExplanations.value][row] = BG_GREEN
#                   ElseIf intAusgabeAuswertungTUs = 2 Then
                elif te_evaluations.int_output == 2:
#                       'Gelb
#                       .Interior.Color = RGB(255, 255, 102)
                    bg_colours[TDAttribute.OperativeTEComparisonTDTC.value][row] = BG_YELLOW
                    bg_colours[TDAttribute.ComparisonExplanations.value][row] = BG_YELLOW
#                   ElseIf intAusgabeAuswertungTUs = 3 Then
                elif te_evaluations.int_output == 3:
#                       'Rot
#                       .Interior.Color = RGB(255, 51, 0)
                    bg_colours[TDAttribute.OperativeTEComparisonTDTC.value][row] = BG_RED
                    bg_colours[TDAttribute.ComparisonExplanations.value][row] = BG_RED
#                   End If
#               End With
#               With rngTDAttribute(7).Offset(lngDatensatz, 0)
#                   .Value = strAusgabeAuswertungTUsDetails
                row_data[TDAttribute.ComparisonExplanations] = te_evaluations.output_details
#                   If intAusgabeAuswertungTUs = 1 Then
                #implemented above
#                       'Grün
#                       .Interior.Color = RGB(51, 204, 51)
#                   ElseIf intAusgabeAuswertungTUs = 2 Then
#                       'Gelb
#                       .Interior.Color = RGB(255, 255, 102)
#                   ElseIf intAusgabeAuswertungTUs = 3 Then
#                       'Rot
#                       .Interior.Color = RGB(255, 51, 0)
#                   End If
#               End With
#               'Ausgabe Erläuterungen zum Vergleich
#               'Ausgabe fehlende TUs bei TD-AAs
#               rngTDAttribute(8).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeAAs
                row_data[TDAttribute.MissingTEsInTDSafeGuards] = te_evaluations.missing_safe_guards
#               'Ausgabe fehlende TUs bei TFs
#               rngTDAttribute(9).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeTFs
                row_data[TDAttribute.MissingTEsInTCs] = te_evaluations.missing_test_cases
#           
#               'Projektspezifische Ausgabe - MEB21
#               If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
                if self.is_project_specific:
#                   rngTDAttribute(26).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Temp11_Auswahlfeld)
                    row_data[TDProjectAttribute.Temp11SelectionField] = unic_join(CRLF, verification_criterion.requirement_temp11_selection_field)
#               End If
                for field, value in row_data.items():
                    td_output_data[field].append(value)
#           
#           End If)
#           
#       Next Verifikationskriterium
#       
#       'Spaltenbreite anpassen
        ###TODO: format colunm width
#       With wksTD.Cells
#           .Columns.EntireColumn.AutoFit
#           .Rows.AutoFit
#       End With
#       
#       'Projektspezifische Sortierung - MEB21
        ###Done at td_attributes variable definition
#       If strProjekt = "MEB21" Then
#           wksTD.Columns(26).Cut
#           wksTD.Columns(14).Insert shift:=xlToRight
        ### Implemented at td_data_output creation
#       End If
#       
#       'Filterung aktivieren
        ###TODO Set filter
#       wksTD.Rows(1).AutoFilter
#       
#       'Dateinamen ausgeben und verstecken
        ###Not implemented
#       lngDatensatz = 1
#       wksTD.Rows(lngDatensatz).EntireRow.Insert shift:=xlDown
#       wksTD.Cells(lngDatensatz, 1) = "Anforderungen: " & strDateinamen(1) & vbCrLf & "Verifikationskriterien: " & strDateinamen(2) & vbCrLf & _
#                                       "Absicherungsaufträge: " & strDateinamen(3) & vbCrLf & "Testfälle: " & strDateinamen(3) & vbCrLf & _
#                                       "FRU-Timing: " & strDateinamen(5)
#       wksTD.Rows(lngDatensatz).EntireRow.Hidden = True
        
        #self.td_output_data = td_output_data
        #self.output_workbook.append_worksheet(
        #    data_frame=pd.DataFrame(
        #        td_output_data,
        #        dtype=str,
        #    ),
        #    name=worksheet_name,
        #)
        return td_output_data, worksheet_name, bg_colours
#       End Sub
#
        ###Moved to test_environment_evaluation ---------------------------------       
#       Private Sub AuswertungTUAbgleich(ByRef intAbgleichTUs() As Integer, ByRef strAbgleichTUs() As String, ByRef intAuswertungTUs() As Integer, ByRef strAuswertungTUs() As String)
#       Private Sub AusgabeTUAbgleich(ByRef intAuswertungTUs() As Integer, ByRef strAuswertungTUs() As String, _
#                                            ByRef strAuswertungTUsFehlendeAAs As String, ByRef strAuswertungTUsFehlendeTFs As String, _
#                                            ByRef intAusgabeAuswertungTUs As Integer, ByRef strAusgabeAuswertungTUs As String, ByRef strAusgabeAuswertungTUsDetails As String)
        ###----------------------------------------------------------------------
#   

#   
#   Private Function AuswertungUnterschiedlicheIStufen(ByRef IStufen As Collection) As Boolean
    def different_i_levels(self, i_levels:list|tuple) -> bool:
#       Dim i As Integer
#       
#       AuswertungUnterschiedlicheIStufen = False
#       
#       If IStufen.Count > 1 Then
#           For i = 2 To IStufen.Count
        for i in range(1, len(i_levels)):
#               If IStufen.Item(i) <> IStufen.Item(i - 1) Then
            if i_levels[i] != i_levels[i-1]:
                return True
#                   AuswertungUnterschiedlicheIStufen = True
#               End If
#           Next i
#       End If
        return False
#   End Function
#   
#   Private Sub AusgabeVerlauf(ByRef wksStatus As Worksheet, ByRef strFehlerVerlauf As String, ByVal intAuswahl As Integer)
    def output_history(self):
        # loop will be use to implement next two calls from the old code
        #               Call AusgabeVerlauf(wksBsM, strFehlerATEVerlauf, 1)
        #               Call AusgabeVerlauf(wksTD, strFehlerTDVerlauf, 2)
        pass #TODO
#       Dim wksVerlauf As Worksheet
#       Dim rngAttributeVerlauf() As Range
#       Dim intAttributeZaehler As Integer
#       Dim intAttributeStatus As Integer
#       Dim blnVerlaufAttribute As Boolean
#       Dim lngVerlaufLetzteZeile As Long
#       Dim lngStatusLetzteZeile As Long
#       Dim intKW As Integer
#       Dim rngVerlaufKWVorhanden As Range
#       
#       'Fehlermeldungen ausschalten
#       On Error Resume Next
#       
#       blnVerlaufAttribute = False
#       strFehlerVerlauf = ""
#       
#       'Arbeitsblatt ATE/TD_Status_Verlauf vorhanden?
#       Select Case intAuswahl:
#           Case 1:
#               Set wksVerlauf = ThisWorkbook.Sheets("ATE_Status_Verlauf")
#           Case 2:
#               Set wksVerlauf = ThisWorkbook.Sheets("TD_Status_Verlauf")
#       End Select
#       
#       'Arbeitsblatt ATE/TD_Status_Verlauf vorhanden
#       If Not wksVerlauf Is Nothing Then
#           blnVerlaufAttribute = True
#           
#           'Attribute aus ATE/TD_Status in ATE/TD_Status_Verlauf suchen
#           intAttributeStatus = wksStatus.Cells(2, 1).End(xlToRight).Column
#       
#           ReDim rngAttributeVerlauf(1 To intAttributeStatus)
#           
#           For intAttributeZaehler = 1 To intAttributeStatus
#               Set rngAttributeVerlauf(intAttributeZaehler) = wksVerlauf.Cells.Find(wksStatus.Cells(2, intAttributeZaehler).Value, lookat:=xlWhole)
#               
#               'Festhalten des Indexes für Attribut "KW Datenauswertung"
#               If wksStatus.Cells(2, intAttributeZaehler).Value = "KW Datenauswertung" Then
#                   intKW = intAttributeZaehler
#               End If
#               
#               'Flag setzen, falls nicht alle Attribute aus ATE/TD_Status in ATE/TD_Status_Verlauf gefunden
#               If rngAttributeVerlauf(intAttributeZaehler) Is Nothing Then
#                   blnVerlaufAttribute = False
#               End If
#           Next intAttributeZaehler
#           
#           'Prüfung, ob KW der Auswertung bereits im Verlauf vorhanden:
#           'KW-Eintrag vorhanden: Abbruch mit Fehlermeldung
#           'KW-Eintrag nicht vorhanden: Werte übernehmen
#           With wksVerlauf.Columns(rngAttributeVerlauf(intKW).Column)
#               Set rngVerlaufKWVorhanden = .Cells.Find(wksStatus.Cells(3, intKW), lookat:=xlWhole)
#           End With
#           
#           If Not rngVerlaufKWVorhanden Is Nothing Then
#           Select Case intAuswahl:
#           Case 1:
#               strFehlerVerlauf = "Eintrag für Kalenderwoche bereits im ATE_Status_Verlauf vorhanden."
#           Case 2:
#               strFehlerVerlauf = "Eintrag für Kalenderwoche bereits im TD_Status_Verlauf vorhanden."
#           End Select
#           
#           Else
#       
#               'Werte von ATE/TD_Status in ATE/TD_Status_Verlauf übernehmen
#               If blnVerlaufAttribute = True Then
#                   'Letzte Zeile im ATE/TD_Status_Verlauf ermitteln
#                   If wksVerlauf.Cells(2, 1).Value <> "" Then
#                       lngVerlaufLetzteZeile = wksVerlauf.Cells(1, 1).End(xlDown).Row
#                   Else
#                       lngVerlaufLetzteZeile = 1
#                   End If
#                   
#                   'Letzte Zeile im ATE_Status ermitteln
#                   If wksStatus.Cells(3, 1).Value <> "" Then
#                       lngStatusLetzteZeile = wksStatus.Cells(2, 1).End(xlDown).Row
#                   Else
#                       lngStatusLetzteZeile = 2
#                   End If
#           
#                   'Falls Werte vorhanden, Werte spaltenweise übernehmen
#                   If lngStatusLetzteZeile > 2 Then
#                       For intAttributeZaehler = 1 To intAttributeStatus
#                           wksStatus.Range(wksStatus.Cells(3, intAttributeZaehler), wksStatus.Cells(lngStatusLetzteZeile, intAttributeZaehler)).Copy _
#                           Destination:=wksVerlauf.Range(wksVerlauf.Cells(lngVerlaufLetzteZeile + 1, rngAttributeVerlauf(intAttributeZaehler).Column), wksVerlauf.Cells(lngVerlaufLetzteZeile + lngStatusLetzteZeile - 2, rngAttributeVerlauf(intAttributeZaehler).Column))
#                       Next intAttributeZaehler
#                   End If
#               End If
#           
#           End If
#       
#       Else
#           Select Case intAuswahl:
#           Case 1:
#               strFehlerVerlauf = "Arbeitsblatt 'ATE_Status_Verlauf' nicht vorhanden"
#           Case 2:
#               strFehlerVerlauf = "Arbeitsblatt 'TD_Status_Verlauf' nicht vorhanden"
#           End Select
#       End If
#       
#   End Sub
#   