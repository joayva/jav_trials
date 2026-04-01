from xls_management.ate.data_de import (
    KNOWN_TEST_ENVIRONMENTS,
    RELEVANT_TEST_ENVIRONMENT_TOP as RELEVANT_TOP,
)

class TestEnvironmentEvaluation:
    def __init__(self):
#       Dim intAuswertungTUs() As Integer               'Integer-Array für die Ergebnisse des Tu-Abgleichs
        self.value = 0
#       Dim strAuswertungTUs() As String                'String-Array für die Ergebnisse des Tu-Abgleichs
        self.test_environments: str = ''

class TestEnvironmentEvaluations:
    indexes = (
        1,   #TF nicht vorhanden, VK fachlich abgestimmt
        10,  #TF operativ, VK nicht vorhanden 
        11,  #TF operativ, VK fachlich abgestimmt 
        20,  #TF nicht operativ, VK nicht vorhanden 
        21,  #TF nicht operativ, VK fachlich abgestimmt 
        30,  #TF teilweise operativ und nicht operativ, VK nicht vorhanden 
        31,  #TF teilweise operativ und nicht operativ, VK fachlich abgestimmt
    )

    def __init__(
            self,
            comparison_count:list[int],
            test_environment_types:list[str] = KNOWN_TEST_ENVIRONMENTS[:RELEVANT_TOP],
        ):
        self.evaluations:dict[int,TestEnvironmentEvaluation] = { 
            i:TestEnvironmentEvaluation() for i in TestEnvironmentEvaluations.indexes
        }
#       Dim strAusgabeAuswertungTUs As String           'String für Ausgabe des Tu-Abgleichs
        self.output = ''
#       Dim intAusgabeAuswertungTUs As Integer          'Integer für Ausgabe des Tu-Abgleichs
        self.int_output = 0
#       Dim strAuswertungTUsFehlendeAAs As String       'String für Ausgabe der fehlenden TUs bei TD-AAs
        self.missing_safe_guards = ''
#       Dim strAuswertungTUsFehlendeTFs As String       'String für Ausgabe der fehlenden TUs bei TFs
        self.missing_test_cases = ''
#       Dim strAusgabeAuswertungTUsDetails As String    'String für Ausgabe des Tu-Abgleichs mit Details
        self.output_details = ''
        #strAbgleichTUs
        self.test_environment_types = test_environment_types
        self.comparison_count = comparison_count

    def summarize(self):
#       Dim i As Integer        'Laufvariable
#       
#       For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
        for index, test_environment_type in enumerate(self.test_environment_types):
            te_index = self.comparison_count[index]
            if te_index in TestEnvironmentEvaluations.indexes:
                self.evaluations[te_index].value += 1
                if self.evaluations[te_index].test_environments == '':
                    self.evaluations[te_index].test_environments = test_environment_type
                else:
                    self.evaluations[te_index].test_environments += f', {test_environment_type}'
#           If intAbgleichTUs(i) = 1 Then
#               '#TF nicht vorhanden, VK fachlich abgestimmt
#                   intAuswertungTUs(1) = intAuswertungTUs(1) + 1
#               If strAuswertungTUs(1) = "" Then
#                   strAuswertungTUs(1) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(1) = strAuswertungTUs(1) & ", " & strAbgleichTUs(i)
#               End If
#               
#           ElseIf intAbgleichTUs(i) = 10 Then
#               '#TF operativ, VK nicht vorhanden
#               intAuswertungTUs(10) = intAuswertungTUs(10) + 1
#               If strAuswertungTUs(10) = "" Then
#                   strAuswertungTUs(10) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(10) = strAuswertungTUs(10) & ", " & strAbgleichTUs(i)
#               End If
#           ElseIf intAbgleichTUs(i) = 11 Then
#               '#TF operativ, VK fachlich abgestimmt
#               intAuswertungTUs(11) = intAuswertungTUs(11) + 1
#               If strAuswertungTUs(11) = "" Then
#                   strAuswertungTUs(11) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(11) = strAuswertungTUs(11) & ", " & strAbgleichTUs(i)
#               End If
#       
#           ElseIf intAbgleichTUs(i) = 20 Then
#               '#TF nicht operativ, VK nicht vorhanden
#               intAuswertungTUs(20) = intAuswertungTUs(20) + 1
#               If strAuswertungTUs(20) = "" Then
#                   strAuswertungTUs(20) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(20) = strAuswertungTUs(20) & ", " & strAbgleichTUs(i)
#               End If
#           ElseIf intAbgleichTUs(i) = 21 Then
#               '#TF nicht operativ, VK fachlich abgestimmt
#               intAuswertungTUs(21) = intAuswertungTUs(21) + 1
#               If strAuswertungTUs(21) = "" Then
#                   strAuswertungTUs(21) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(21) = strAuswertungTUs(21) & ", " & strAbgleichTUs(i)
#               End If
#               
#           ElseIf intAbgleichTUs(i) = 30 Then
#               '#TF teilweise operativ und nicht operativ, VK nicht vorhanden
#               intAuswertungTUs(30) = intAuswertungTUs(30) + 1
#               If strAuswertungTUs(30) = "" Then
#                   strAuswertungTUs(30) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(30) = strAuswertungTUs(30) & ", " & strAbgleichTUs(i)
#               End If
#           ElseIf intAbgleichTUs(i) = 31 Then
#               '#TF teilweise operativ und nicht operativ, VK fachlich abgestimmt
#               intAuswertungTUs(31) = intAuswertungTUs(31) + 1
#               If strAuswertungTUs(31) = "" Then
#                   strAuswertungTUs(31) = strAbgleichTUs(i)
#               Else
#                   strAuswertungTUs(31) = strAuswertungTUs(31) & ", " & strAbgleichTUs(i)
#               End If
#           End If
#       Next i
#   End Sub

#   Private Sub AusgabeTUAbgleich(ByRef intAuswertungTUs() As Integer, ByRef strAuswertungTUs() As String, _
#                                            ByRef strAuswertungTUsFehlendeAAs As String, ByRef strAuswertungTUsFehlendeTFs As String, _
#                                            ByRef intAusgabeAuswertungTUs As Integer, ByRef strAusgabeAuswertungTUs As String, ByRef strAusgabeAuswertungTUsDetails As String)
    def output_comparison(self): 
#       '1) alle TF operativ und VK:TU = TF:TU
#       '2) TF vorhanden, aber Status != operativ oder VK:TU != TF:TU
#       '3) keine TF vorhanden
#       
#       'VKs sind obligatorisch
#       If intAuswertungTUs(1) = 0 Then
        if self.evaluations[1].value == 0:
#           'Keine relevanten VK-TUs ohne TF vorhanden
#           If intAuswertungTUs(11) > 0 Then
            if self.evaluations[11].value > 0:
#               'Alle relevanten VK-TUs mit TF (operativ) abgedeckt
#               strAusgabeAuswertungTUs = "Alle relevanten Testumgebungstypen abgedeckt"
                self.output = 'Alle relevanten Testumgebungstypen abgedeckt'
#               'strAusgabeAuswertungTUsDetails = "Alle relevanten Testumgebungstypen mit operativen Testfällen abgedeckt"
#               strAusgabeAuswertungTUsDetails = "Testfälle vollständig"
                self.output_details = 'Testfälle vollständig'
#               intAusgabeAuswertungTUs = 1
                self.int_output = 1
#           End If
#           If intAuswertungTUs(21) > 0 Then
            if self.evaluations[21].value > 0:
#               'Alle relevanten VK-TUs mit TF (nicht operativ) abgedeckt
#               strAusgabeAuswertungTUs = "Alle relevanten Testumgebungstypen abgedeckt"
                self.output = 'Alle relevanten Testumgebungstypen abgedeckt'
#               'strAusgabeAuswertungTUsDetails = "Alle relevanten Testumgebungstypen mit nicht operativen Testfällen abgedeckt"
#               strAusgabeAuswertungTUsDetails = "Testfälle unvollständig
                self.output_details = 'Testfälle unvollständig'
#               intAusgabeAuswertungTUs = 2
                self.int_output = 2
#           End If
#           If intAuswertungTUs(31) > 0 Then
            if  self.evaluations[31].value > 0:
#               'Alle relevanten VK-TUs mit TF (teilweise operativ) abgedeckt
#               strAusgabeAuswertungTUs = "Alle relevanten Testumgebungstypen abgedeckt"
                self.output = 'Alle relevanten Testumgebungstypen abgedeckt'
#               'strAusgabeAuswertungTUsDetails = "Alle relevanten Testumgebungstypen mit operativen und nicht operativen Testfällen abgedeckt"
#               strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
                self.output_details = 'Testfälle unvollständig'
#               intAusgabeAuswertungTUs = 2
                self.int_output = 2
#           End If
#           If intAuswertungTUs(11) = 0 And intAuswertungTUs(21) = 0 And intAuswertungTUs(31) = 0 Then
            if(
                self.evaluations[11].value == 0 and 
                self.evaluations[21].value == 0 and 
                self.evaluations[31].value == 0
            ):
#               'Keine relevanten VK-TUs vorhanden
#               strAusgabeAuswertungTUs = "Keine relevanten Testumgebungstypen vorhanden"
                self.output = 'Keine relevanten Testumgebungstypen vorhanden'
#               intAusgabeAuswertungTUs = 1
                self.int_output = 1
#           End If
#           'Feldfarbe grün
#       ElseIf intAuswertungTUs(1) > 0 Then
        else: # self.evaluations[1].value > 0
#           'Relevante VK-TUs ohne TF vorhanden
#           If intAuswertungTUs(11) > 0 Then
            if self.evaluations[11].value > 0:
#               'Einige relevanten VK-TUs mit TF (operativ) abgedeckt
#               strAusgabeAuswertungTUs = "Relevante Testumgebungstypen teilweise abgedeckt"
                self.output = 'Relevante Testumgebungstypen teilweise abgedeckt'
#               'strAusgabeAuswertungTUsDetails = "Relevante Testumgebungstypen teilweise mit operativen Testfällen abgedeckt"
#               strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
                self.output_details = 'Testfälle unvollständig'
#               'Feldfarbe gelb
#               intAusgabeAuswertungTUs = 2
                self.int_output = 2
#           End If
#           If intAuswertungTUs(21) > 0 Then
            if self.evaluations[21].value > 0:
#               'Einige relevanten VK-TUs mit TF (nicht operativ) abgedeckt
#               strAusgabeAuswertungTUs = "Relevante Testumgebungstypen teilweise abgedeckt"
                self.output = 'Relevante Testumgebungstypen teilweise abgedeckt'
#               'strAusgabeAuswertungTUsDetails = "Relevante Testumgebungstypen teilweise mit nicht operativen Testfällen abgedeckt"
#               strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
                self.output_details = 'Testfälle unvollständig'
#               'Feldfarbe gelb
#               intAusgabeAuswertungTUs = 2
                self.int_output = 2
#           End If
#           If intAuswertungTUs(31) > 0 Then
            if self.evaluations[31].value > 0:
#               'Einige relevanten VK-TUs mit TF (teilweise operativ) abgedeckt
#               strAusgabeAuswertungTUs = "Relevante Testumgebungstypen teilweise abgedeckt"
                self.output = 'Relevante Testumgebungstypen teilweise abgedeckt'
#               'strAusgabeAuswertungTUsDetails = "Relevante Testumgebungstypen teilweise mit operativen und nicht operativen Testfällen abgedeckt"
#               strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
                self.output_details = 'Testfälle unvollständig'
#               'Feldfarbe gelb
#               intAusgabeAuswertungTUs = 2
                self.int_output = 2
#           End If
#           If intAuswertungTUs(11) = 0 And intAuswertungTUs(21) = 0 And intAuswertungTUs(31) = 0 Then
            if(
                self.evaluations[11].value == 0 and 
                self.evaluations[21].value == 0 and 
                self.evaluations[31].value == 0
            ):
#               'Keine relevanten VK-TUs abgedeckt
#               strAusgabeAuswertungTUs = "Relevante Testumgebungstypen nicht abgedeckt"
                self.output = 'Relevante Testumgebungstypen nicht abgedeckt'
#               strAusgabeAuswertungTUsDetails = "Keine Testfälle vorhanden"
                self.output_details = 'Keine Testfälle vorhanden'
#               'Feldfarbe rot
#               intAusgabeAuswertungTUs = 3
                self.int_output = 3
#           End If
#           'Fehlende TUs bei TFs erfassen
#           strAuswertungTUsFehlendeTFs = strAuswertungTUs(1)
            self.missing_test_cases = self.evaluations[1].test_environments
#       End If
#       
#       'Weitere relevante TUs in TFs vorhanden?
#       If intAuswertungTUs(10) > 0 Then
        if self.evaluations[10].value > 0:
#           'TF operativ
#           'strAusgabeAuswertungTUsDetails = strAusgabeAuswertungTUsDetails & vbCrLf & "Weitere operative Testfälle für abweichende Testumgebungstypen vorhanden."
#           strAuswertungTUsFehlendeAAs = strAuswertungTUs(10)
            self.missing_safe_guards =  self.evaluations[10].test_environments
#       End If
#       If intAuswertungTUs(20) > 0 Then
        if self.evaluations[20].value > 0:
#           'TF nicht operativ
#           'strAusgabeAuswertungTUsDetails = strAusgabeAuswertungTUsDetails & vbCrLf & "Weitere nicht operative Testfälle für abweichende Testumgebungstypen vorhanden."
#           If strAuswertungTUsFehlendeAAs = "" Then
            if self.missing_safe_guards == '':
#               strAuswertungTUsFehlendeAAs = strAuswertungTUs(20)
                self.missing_safe_guards = self.evaluations[20].test_environments
#           Else
            else:
#               strAuswertungTUsFehlendeAAs = strAuswertungTUsFehlendeAAs & ", " & strAuswertungTUs(20)
                self.missing_safe_guards = f'{self.missing_safe_guards}, {self.evaluations[20].test_environments}'
#           End If
#       End If
#       If intAuswertungTUs(30) > 0 Then
        if self.evaluations[30].value > 0:
#           'TF operativ und nicht operativ
#           'strAusgabeAuswertungTUsDetails = strAusgabeAuswertungTUsDetails & vbCrLf & "Weitere operative und nicht operative Testfälle für abweichende Testumgebungstypen vorhanden."
#           If strAuswertungTUsFehlendeAAs = "" Then
#               strAuswertungTUsFehlendeAAs = strAuswertungTUs(30)
#           Else
#               strAuswertungTUsFehlendeAAs = strAuswertungTUsFehlendeAAs & ", " & strAuswertungTUs(30)
#           End If
#           strAuswertungTUsFehlendeAAs = strAuswertungTUs(30)
            # This sentence overwrites value set in If bock above
            self.missing_safe_guards = self.evaluations[30].test_environments
#       End If
#   End Sub

    def empty(self, output:str):       
#       '<output>
#       strAusgabeAuswertungTUs = <output>
        self.output = output
#       intAusgabeAuswertungTUs = 3
        self.int_output = 3
