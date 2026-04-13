
from pathlib import Path

import pandas as pd

from xls_management.ate.om.db_info import DBInfo
from xls_management.tui.file_picker import path_from_file_picker
from xls_management.tui.yes_no_form import yes_no_msgbox
from xls_management.xlsx.workbook import Workbook


class ProjectDBInfo(DBInfo):
    def __init__(
        self,
        **args
    ):
        """
        Project specific DBInfo
        
        :param self: ProjectDBInfo object reference
        :param args: Next arguments are expected         
                        db_info:DBInfo ,(optional) DBInfo object; attributes=db_info.attributes
                        attributes: tuple[str, ...], (optional) should be provided if no db_info object
                        project:str,
                        project_attributes: tuple[str, ...]= (),
        """
        attributes:tuple[str, ...]
        if 'db_info' in args.keys():
            attributes=args['db_info'].attributes
        else:
            assert 'attributes' in args.keys(), "attributes param should be provided"
            attributes=args['attributes']
        super().__init__(attributes=attributes)
        assert 'project' in args.keys(),"project param should be provided"
        self.project:str = args['project']
        assert 'project_attributes' in args.keys(),"project_attributes param should be provided"
        self.project_attributes:tuple[str, ...] = args['project_attributes']

    
#   Public Function EinlesenDatei_Projektspezifisch(ByVal strTitel As String, ByRef strAttribute() As String, ByRef rngAttribute() As Range, ByRef wbImport As Workbook, ByRef wksImport As Worksheet, ByRef strFehler As String, ByRef strDateinamen As String, _
#                                                       ByVal strProjekt As String, ByRef strAttributeProjekt() As String, ByRef rngAttributeProjekt() As Range) As Boolean
    def read_file(self, titel:str, import_file_path:str='') -> bool:
        """
        user chooses a workbook using a file picker widget
        True,"" is returned if each expected attribute is in one of the workbook sheets;
                self.sheet_name is set with the name of the sheet containg those attributes
        False, error_trace is returned elsewhere; 
                being error trace a trace of missing attributes
           
        """ 
#       Dim blnAttributeImport As Boolean   'Flag für Attribute
#       Dim blnAttributeProjektImport As Boolean    'Flag für projektspezifische Attribute
#       Dim intWksImport As Integer         'Zähler für Worksheets
#       Dim strImportPfad As String         'String für Dateipfad
#       Dim strImportDatei As String        'String für Dateinamen
#       Dim strFehlerGesamt As String       'Sting für Gesamtfehler
#       Dim intAttributeZaehler As Integer  'Zähler für Attribute
#       Dim strFehlerAttribute As String    'String für fehlende Attribute
#       Dim intAttributeProjektZaehler As Integer   'Zähler für Projekt-Attribute
#       Dim strFehlerProjektAttribute As String 'String für fehlende Projekt-Attribute
#       Dim blnErsterFund As Boolean        'Flag für erstes Auffinden des vollständigen Attributesatzes
#       Dim rngAttributeErsterFund() As Range   'Range-Array für erstes Auffinden des vollständigen allgemeinen Attributesatzes
#       Dim rngAttributeProjektErsterFund() As Range    'Range-Array für erstes Auffinden des vollständigen projektspezifischen Attributesatzes
#       Dim wksImportErsterFund As Worksheet    'Worksheet für erstes Auffinden des vollständigen Attributesatzes
#       Dim lngZeilenZahlerErsterFund As Long   'Long für letzte Zeile des ersten Fundes
#       Dim lngWeitererFundZaehler As Long  'Long für Startzeile des zusätzlichen Datensatzen
#       
#       On Error Resume Next
#       
#       blnAttributeImport = False
        attribute_import = False
#       blnAttributeProjektImport = False
        project_attribute_import = False
#       strFehlerGesamt = ""
        self.error_msg = ""
#       strDateinamen = ""
#       blnErsterFund = False
        first_find = False
        first_find_columns:pd.DataFrame|None = None
#       
#       strImportPfad = Application.GetOpenFilename(FileFilter:="Excel-Dateien (*.xls; *.xlsx; *.xlm; *.xlsm), *.xls; *.xlsx; *.xlm; *.xlsm", FilterIndex:=1, Title:=strTitel & " auswählen")
        if import_file_path == '':
            import_file_path = path_from_file_picker(location=".", title= f"{titel} should be chosen")
        elif isinstance(import_file_path,str):
            import_file_path = Path(import_file_path)
#       If Trim(strImportPfad) <> "Falsch" Then
        if import_file_path is not None:
#           'Datei öffnen
#           Workbooks.Open Filename:=strImportPfad, ReadOnly:=True
            workbook: Workbook = Workbook(import_file_path)
#           'Dateinamen extrahieren
#           strImportDatei = Right(strImportPfad, Len(strImportPfad) - InStrRev(strImportPfad, "\"))
#           strDateinamen = strImportDatei
            import_file_name = Path(workbook.file_path).name
#           'Workbook zuweisen
#           Set wbImport = Workbooks(strImportDatei)
#           'Worksheets nach Attributen durchsuchen
#           intWksImport = 0
#           Do
            for self.sheet_name, self.columns in workbook.all_sheets():
#               'Rangeobjekte zurücksetzen
#               ReDim rngAttribute(LBound(strAttribute, 1) To UBound(strAttribute, 1))
#               ReDim rngAttributeProjekt(LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1))
#               
#               'Zähler für Arbeitsblatt erhöhen
#               intWksImport = intWksImport + 1
#               'Worksheet zuweisen
#               Set wksImport = wbImport.Sheets(intWksImport)
#               
#               'Allgemeine Attribute suchen
#               For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                   Set rngAttribute(intAttributeZaehler) = wksImport.Cells.Find(strAttribute(intAttributeZaehler), lookat:=xlWhole)
#               Next intAttributeZaehler
                missing_attributes = [attribute for attribute in self.attributes if attribute not in self.columns]
#               
#               'Projekt-Attribute suchen
#               For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                   Set rngAttributeProjekt(intAttributeProjektZaehler) = wksImport.Cells.Find(strAttributeProjekt(intAttributeProjektZaehler), lookat:=xlWhole)
#               Next intAttributeProjektZaehler
                missing_project_attributes = [attribute for attribute in self.project_attributes if attribute not in self.columns]
#               
#               'Suche der allgemeinen Attribute auswerten
#               strFehlerAttribute = ""
#               If RangeObjekteVorhandenFehlerausgabe(rngAttribute, strAttribute, strFehlerAttribute) Then
                if len(missing_attributes) == 0:
                    # all expected attributes have been found; self.sheet_name and self.columns are set
                    attribute_import = True
#                   blnAttributeImport = True
#               Else
#                   If strFehlerGesamt = "" Then
#                       strFehlerGesamt = wksImport.Name & ": " & strFehlerAttribute
#                   Else
#                       strFehlerGesamt = strFehlerGesamt & "; " & wksImport.Name & ": " & strFehlerAttribute
#                   End If
#               End If
                else:
                    self.trace_error(import_file_name, missing_attributes)
#               
#               'Suche der Projekt-Attribute auswerten
#               strFehlerProjektAttribute = ""
#               If RangeObjekteVorhandenFehlerausgabe(rngAttributeProjekt, strAttributeProjekt, strFehlerProjektAttribute) Then
#                   blnAttributeProjektImport = True
                if len(missing_project_attributes) == 0:
                    project_attribute_import = True
#               Else
#                   If strFehlerGesamt = "" Then
#                       strFehlerGesamt = wksImport.Name & ": " & strProjekt & " - " & strFehlerProjektAttribute
#                   Else
#                       strFehlerGesamt = strFehlerGesamt & "; " & wksImport.Name & ": " & strProjekt & " - " & strFehlerProjektAttribute
#                   End If
#               End If
                else:
                    self.trace_project_error(import_file_name, missing_project_attributes)
#               
#               'Auswertung beider Suchen (allgemein und projektspezifisch) und Unterscheidung, ob es das erste Auffinden des vollständigen Attributsatzes ist oder weitere
#               'Beim ersten Auffinden werden die beiden Range-Arays gespeichert
#               'Ab dem zweitem Auffinden des vollständigen Attributsatzes erfolgt eine Abfrage, eine Kopiervorganges der dazugehörigen Daten zum ersten Datensatz erfolgen soll und ggf. Umsetzung
#               If blnAttributeImport = True And blnAttributeProjektImport = True Then
                if attribute_import and project_attribute_import:
#                   'Zurücksetzen der Flags für den einzelnen Durchlauf
#                   blnAttributeImport = False
                    attribute_import = False
#                   blnAttributeProjektImport = False
                    project_attribute_import = False
#                   
#                   If blnErsterFund = False Then
                    if first_find_columns is None:
#                       blnErsterFund = True
                        first_find = True
#                       'Felder dimensionieren
#                       ReDim rngAttributeErsterFund(LBound(strAttribute, 1) To UBound(strAttribute, 1))
#                       ReDim rngAttributeProjektErsterFund(LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1))
#                       'Übernahme allgemeine Attribute
#                       For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                           Set rngAttributeErsterFund(intAttributeZaehler) = rngAttribute(intAttributeZaehler)
#                       Next intAttributeZaehler
#                       'Übernahme projektspezifische Attribute
#                       For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                           Set rngAttributeProjektErsterFund(intAttributeProjektZaehler) = rngAttributeProjekt(intAttributeProjektZaehler)
#                       Next intAttributeProjektZaehler
#                       'Übernahme Arbeitsblatt
#                       Set wksImportErsterFund = wksImport
                        first_find_columns = self.columns
                        first_find_sheet_name = self.sheet_name
                        
                        
#                   Else
#                       If MsgBox("Weiterer Datensatz gefunden!" & vbCrLf & "Erster Fund: " & wksImportErsterFund.Name & vbCrLf & "Aktueller Fund: " & wksImport.Name & vbCrLf & vbCrLf & "Sollen die Datensätze zusammengeführt werden?", vbYesNo) = vbYes Then
                    elif yes_no_msgbox(
                        f"More data found!\nFirst found: {first_find_sheet_name}\n"
                        f"Current found:{self.sheet_name}\n"
                        "Should the data sets be merged?"
                    ):
                        # append self.columns to first_find_columns
                        first_find_columns = pd.concat([first_find_columns, self.columns], ignore_index=True)
#                       End If
#                   End If
#               End If
#               
#           'Loop While intWksImport < wbImport.Worksheets.Count And blnAttributeImport = False And blnAttributeProjektImport = False
#           Loop While intWksImport < wbImport.Worksheets.Count
#       End If
#       
#       
#       'Rückgabe allgemeine Attribute
#       For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#           Set rngAttribute(intAttributeZaehler) = rngAttributeErsterFund(intAttributeZaehler)
#       Next intAttributeZaehler
#       'Rückgabe projektspezifische Attribute
#       For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#           Set rngAttributeProjekt(intAttributeProjektZaehler) = rngAttributeProjektErsterFund(intAttributeProjektZaehler)
#       Next intAttributeProjektZaehler
#       'Rückgabe Arbeitsblatt
#       Set wksImport = wksImportErsterFund
        self.columns = first_find_columns
#       
#       'Rückgabewert übernehmen
#       EinlesenDatei_Projektspezifisch = blnErsterFund
#       'Fehlerwert übernehmen
#       strFehler = strFehlerGesamt
        return first_find
#   End Function

    #def append_current_to_first_found(self, columns:pd.DataFrame, first_find_columns:pd.DataFrame):
    #    pass
    #    #TOBEDEL not required, pd.concat was used instead
#      'Alle Zeilen mit Einträgen durchlaufen und kopieren
#      For lngWeitererFundZaehler = 1 To wksImport.UsedRange.Rows.Count - rngAttribute(1).Row
#          If rngAttribute(1).Offset(lngWeitererFundZaehler, 0).Value <> "" Then
#              'Zielzeile ermitteln
#              lngZeilenZahlerErsterFund = rngAttributeErsterFund(1).End(xlDown).Row + 1
#              
#              'Allgemeine Werte kopieren
#              For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                  wksImportErsterFund.Cells(lngZeilenZahlerErsterFund, rngAttributeErsterFund(intAttributeZaehler).Column).Value = rngAttribute(intAttributeZaehler).Offset(lngWeitererFundZaehler).Value
#              Next intAttributeZaehler
#              
#              'Projektspezifische Werte kopieren
#              For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                  wksImportErsterFund.Cells(lngZeilenZahlerErsterFund, rngAttributeProjektErsterFund(intAttributeProjektZaehler).Column).Value = rngAttributeProjekt(intAttributeProjektZaehler).Offset(lngWeitererFundZaehler).Value
#              Next intAttributeProjektZaehler
#          End If
#      Next lngWeitererFundZaehler


    def trace_project_error(self, import_file_name, missing_attributes) -> None:
        self.error_msg += (
                        f"The following {self.project} project attributes are missing from {self.sheet_name} in {import_file_name}: "
                        f"{', '.join(missing_attributes)}\n"
                    )
