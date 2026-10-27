VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main 
   Caption         =   "DataCo Global - Gestion logistique"
   ClientHeight    =   11460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16560
   OleObjectBlob   =   "new-design-F_Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRefreshing As Boolean
Private mReady As Boolean
Private mLayoutBusy As Boolean
Private mLayoutError As String
Private WithEvents mProductSelector As MSForms.ComboBox
Attribute mProductSelector.VB_VarHelpID = -1
Private WithEvents mReportSelector As MSForms.ComboBox
Attribute mReportSelector.VB_VarHelpID = -1
Private mReportBusy As Boolean
Private mReportWidths As String
Private mReportHeaders As Variant
Private mDesignButtons As Collection
Private mContentZoom As Long
Private mPageZoom(0 To 10) As Long

Private Sub UserForm_Initialize()
    On Error GoTo Failed
    EnsureWorkbookVisible
    mRefreshing = True
    pr_Initialize
    fu_Initialize
    co_Initialize
    ap_Initialize
    st_Initialize
    par13.AddItem "Oui": par13.AddItem "Non"
    InstallProductSelector
    InstallReportViewer
    InstallDesignControls
    LoadSettings
    LoadGuide
    mRefreshing = False
    st_cboProduct_Change
    mReady = True
    RefreshHome
    RefreshResults
    FitWindow
    Exit Sub
Failed:
    mRefreshing = False
    MsgBox "Ouverture de l'interface : " & Err.Description, vbExclamation
End Sub

Private Sub InstallProductSelector()
    Dim p As Object
    Set p = tabsMain.Pages(6)
    p.Controls("par15").Visible = False
    Set mProductSelector = p.Controls.Add("Forms.ComboBox.1", "cboSimulationProduct", True)
    With mProductSelector
        .Move p.Controls("par15").Left, p.Controls("par15").Top, p.Controls("par15").Width, p.Controls("par15").Height
        .ColumnCount = 2
        .BoundColumn = 1
        .TextColumn = 2
        .ColumnWidths = "52 pt;170 pt"
        .Style = fmStyleDropDownList
    End With
    FillProducts mProductSelector
    p.Controls("lblParam15").Caption = "Produit simule (4 choix)"
    ThisWorkbook.Worksheets("Feuille_Controle").Range("D21").Value2 = "Le produit choisi pilote la calibration et les derniers resultats Power BI."
End Sub

Private Sub InstallReportViewer()
    Dim p As Object, c As Object
    Set p = tabsMain.Pages.Add("pgRapports", "Rapports")
    Set c = p.Controls.Add("Forms.ListBox.1", "lstReport", True)
    c.ColumnCount = 8
    c.IntegralHeight = False
    c.MultiSelect = fmMultiSelectSingle
    c.SpecialEffect = 0
    c.BorderStyle = 1
    Set mReportSelector = p.Controls.Add("Forms.ComboBox.1", "cboReportView", True)
    With mReportSelector
        .Style = fmStyleDropDownList
        .Font.name = "Segoe UI": .Font.size = 11
        .AddItem "Synthese des scenarios"
        .AddItem "Evolution quotidienne du stock"
        .AddItem "Disponibilite et decomposition des couts"
        .AddItem "Comparaisons appariees et IC a 95 %"
        .AddItem "Sensibilite : commandes multiples Oui / Non"
        .AddItem "Simulation personnalisee s / Q"
    End With
    mReportBusy = True: mReportSelector.ListIndex = 0: mReportBusy = False
End Sub

Private Sub mProductSelector_Change()
    If mRefreshing Or mProductSelector Is Nothing Then Exit Sub
    If Len(CStr(mProductSelector.value)) = 0 Then Exit Sub
    SetParam "ProduitSimulation", CStr(mProductSelector.value)
    On Error GoTo Failed
    Calibrate True
    LoadSettings
    lblStatus.Caption = "Demande, S0 et suggestions s / Q recalcules pour " & CStr(mProductSelector.list(mProductSelector.ListIndex, 1)) & _
        " (mu=" & Format$(CDbl(Param("mu")), "0.00") & ", sigma=" & Format$(CDbl(Param("sigma")), "0.00") & ")."
    Exit Sub
Failed:
    MsgBox "Calibration du produit : " & Err.Description, vbExclamation
End Sub

Private Sub FitWindow()
    On Error GoTo Failed
    mLayoutBusy = True
    Me.Zoom = 100
    Me.Width = Application.Min(1160, Application.Width - 24)
    Me.Height = Application.Min(760, Application.Height - 24)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    mLayoutBusy = False
    ArrangeInterface
    Exit Sub
Failed:
    mLayoutBusy = False
    mLayoutError = Err.Description
    lblStatus.Caption = "Affichage : " & mLayoutError
End Sub

Private Sub ArrangeInterface()
    If mLayoutBusy Then Exit Sub
    On Error GoTo Failed
    mLayoutBusy = True
    LayoutDataCoForm Me
    ApplyDesignZoom
    PaintDataCoNavigation Me
    If Len(mReportWidths) > 0 Then FitReportColumns tabsMain.Pages("pgRapports"), mReportWidths, mReportHeaders
    mLayoutError = ""
    mLayoutBusy = False
    Exit Sub
Failed:
    mLayoutBusy = False
    mLayoutError = Err.Description
    lblStatus.Caption = "Affichage : " & mLayoutError
End Sub

Private Sub UserForm_Activate()
    If mReady And Not mLayoutBusy Then
        ArrangeInterface
        If tabsMain.value = 10 Then RenderReport
    End If
End Sub

Private Sub UserForm_Resize()
    If mReady And Not mLayoutBusy Then ArrangeInterface
End Sub

Public Sub RefreshLayout()
    ArrangeInterface
    If tabsMain.value = 10 Then RenderReport
End Sub

Public Property Get DisplayError() As String
    DisplayError = mLayoutError
End Property

Public Sub SelectPage(ByVal index As Long)
    tabsMain.value = index
    RefreshPage
End Sub

Private Sub tabsMain_Change()
    If mReady And Not mLayoutBusy Then RefreshPage
End Sub

Private Sub RefreshChoice(ByVal combo As Object, Optional ByVal supplier As Boolean = False)
    Dim oldKey As String, i As Long
    oldKey = CStr(combo.value)
    If supplier Then FillSuppliers combo Else FillProducts combo
    For i = 0 To combo.ListCount - 1
        If CStr(combo.list(i, 0)) = oldKey Then combo.ListIndex = i: Exit For
    Next i
End Sub

Public Sub RefreshPage()
    On Error GoTo Failed
    mRefreshing = True
    Select Case tabsMain.value
        Case 0: RefreshHome
        Case 1: pr_RefreshList
        Case 2: fu_RefreshList
        Case 3: RefreshChoice co_cboProduct: co_RefreshList
        Case 4: RefreshChoice ap_cboProduct: RefreshChoice ap_cboSupplier, True: ap_RefreshList
        Case 5: RefreshChoice st_cboProduct: st_RefreshList
        Case 8: RefreshResults
        Case 10: RenderReport
    End Select
    mRefreshing = False
    If tabsMain.value = 5 Then st_cboProduct_Change
    PaintDataCoNavigation Me
    mContentZoom = mPageZoom(tabsMain.value)
    ZoomDataCoPages Me, mContentZoom
    lblStatus.Caption = tabsMain.SelectedItem.Caption
    Exit Sub
Failed:
    mRefreshing = False
    lblStatus.Caption = Err.Description
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub RefreshHome()
    UpdateOperationalKPI
    With ThisWorkbook.Worksheets("Menu")
        lblCounters.Caption = "Produits sous seuil : " & .Range("E35").Value2 & vbCrLf & "Approvisionnements en attente : " & .Range("E36").Value2
    End With
    lblHomeState.Caption = "Resultats : " & ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2
    If mReady And mContentZoom > 0 And Not mLayoutBusy Then ArrangeInterface
End Sub

Public Sub LoadSettings()
    Dim sh As Worksheet, i As Long, v As Variant, wasRefreshing As Boolean
    wasRefreshing = mRefreshing: mRefreshing = True
    Set sh = ThisWorkbook.Worksheets("Feuille_Controle")
    For i = 0 To 15
        v = sh.Cells(i + 6, 3).Value2
        If i = 12 Then v = CDbl(v) * 100
        tabsMain.Pages(6).Controls("par" & i).value = CStr(v)
    Next i
    mProductSelector.value = CStr(sh.Range("C21").Value2)
    For i = 1 To 3
        With Tbl("tblScenarios").DataBodyRange
            tabsMain.Pages(7).Controls("scenarioName" & i).Caption = .Cells(i, 2).Value2
            tabsMain.Pages(7).Controls("scenarios" & i).value = .Cells(i, 4).Value2
            tabsMain.Pages(7).Controls("scenarioq" & i).value = .Cells(i, 5).Value2
        End With
    Next i
    mRefreshing = wasRefreshing
End Sub

Public Function ApplySettings(Optional ByVal notify As Boolean = False, Optional ByVal showErrors As Boolean = True) As Boolean
    Dim sh As Worksheet, oldValues As Variant, oldScenarios As Variant, v(1 To 16, 1 To 1) As Variant
    Dim policies(1 To 3, 1 To 2) As Variant, i As Long, key As String, changed As Boolean
    Dim events As Boolean, started As Boolean, message As String
    On Error GoTo Failed
    Set sh = ThisWorkbook.Worksheets("Feuille_Controle")
    oldValues = sh.Range("C6:C21").Value2
    oldScenarios = Tbl("tblScenarios").DataBodyRange.columns(4).Resize(3, 2).Value2
    For i = 0 To 15
        key = CStr(sh.Cells(i + 6, 2).Value2)
        If i = 15 Then
            v(i + 1, 1) = CStr(mProductSelector.value)
        ElseIf i = 13 Or i = 14 Then
            v(i + 1, 1) = CStr(tabsMain.Pages(6).Controls("par" & i).value)
        Else
            v(i + 1, 1) = NumberValue(tabsMain.Pages(6).Controls("par" & i).value, key)
        If i = 12 Then v(i + 1, 1) = CDbl(v(i + 1, 1)) / 100
        End If
        If CStr(v(i + 1, 1)) <> CStr(oldValues(i + 1, 1)) Then changed = True
    Next i
    For i = 1 To 3
        policies(i, 1) = NumberValue(tabsMain.Pages(7).Controls("scenarios" & i).value, "Seuil s", 0, True)
        policies(i, 2) = NumberValue(tabsMain.Pages(7).Controls("scenarioq" & i).value, "Quantite Q", 1, True)
        Require policies(i, 1) <= 1000000 And policies(i, 2) <= 1000000, "s et Q doivent etre <= 1000000."
        If policies(i, 1) <> oldScenarios(i, 1) Or policies(i, 2) <> oldScenarios(i, 2) Then changed = True
    Next i
    events = Application.EnableEvents
    Application.EnableEvents = False
    started = True
    sh.Range("C6:C21").Value2 = v
    Tbl("tblScenarios").DataBodyRange.columns(4).Resize(3, 2).Value2 = policies
    ValidateParameters
    SyncScenarioInitialStock
    If changed Then MarkStale
    Application.EnableEvents = events
    started = False
    ApplySettings = True
    lblStatus.Caption = "Parametres et scenarios enregistres."
    If notify Then MsgBox lblStatus.Caption, vbInformation
    Exit Function
Failed:
    message = Err.Description
    If started Then
        sh.Range("C6:C21").Value2 = oldValues
        Tbl("tblScenarios").DataBodyRange.columns(4).Resize(3, 2).Value2 = oldScenarios
        Application.EnableEvents = events
    End If
    lblStatus.Caption = message
    If showErrors Then MsgBox message, vbExclamation
End Function

Private Sub btnSettingsSave_Click()
    Call ApplySettings(True)
End Sub
Private Sub btnScenarioSave_Click()
    Call ApplySettings(True)
End Sub
Private Sub btnSettingsReload_Click()
    LoadSettings
    lblStatus.Caption = "Valeurs rechargees depuis le classeur."
End Sub
Private Sub btnSettingsReset_Click()
    If MsgBox("Reinitialiser les parametres et les trois scenarios ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    ResetParameters
    LoadSettings
End Sub
Private Sub btnCalibrate_Click()
    On Error GoTo Failed
    If Not ApplySettings() Then Exit Sub
    Calibrate True
    LoadSettings
    lblStatus.Caption = "Demande, S0 et suggestions s / Q recalcules pour ce produit."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub btnNormal_Click()
    Dim a As Variant, i As Long
    On Error GoTo Failed
    If Not ApplySettings() Then Exit Sub
    ValidateParameters
    TestNormal
    a = ThisWorkbook.Worksheets("Feuille_Controle").Range("F6:G15").Value2
    lstDraws.Clear
    For i = 1 To 10
        lstDraws.AddItem a(i, 1)
        lstDraws.list(i - 1, 1) = a(i, 2)
    Next i
    lblStatus.Caption = "10 tirages generes."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub btnCompare_Click()
    On Error GoTo Failed
    If Not ApplySettings() Then Exit Sub
    Me.MousePointer = 11
    CompareScenarios True
    Me.MousePointer = 0
    SelectPage 8
    lblStatus.Caption = "Comparaison terminee."
    Exit Sub
Failed:
    Me.MousePointer = 0
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub btnCustom_Click()
    On Error GoTo Failed
    If Not ApplySettings() Then Exit Sub
    RunCustom True
    ShowCustomReport
    lblStatus.Caption = "Simulation terminee pour " & CStr(mProductSelector.value) & ". Ouvrez les resultats personnalises."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub btnCustomResults_Click()
    ShowCustomReport
End Sub

Public Sub RefreshResults()
    Dim a As Variant, out(0 To 2, 0 To 7) As Variant, i As Long, j As Long, cols As Variant, fmt As String
    ThisWorkbook.Worksheets("Synthese").Calculate
    a = ThisWorkbook.Worksheets("Synthese").Range("B8:I10").Value2
    cols = Array(1, 2, 3, 4, 5, 8, 6, 7)
    lblResultState.Caption = "Etat des resultats : " & ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2
    lblResultDecision.Caption = "Decision : " & ThisWorkbook.Worksheets("Synthese").Range("B5").Value2
    For i = 0 To 2
        For j = 0 To 7
            If j = 0 Then
                out(i, j) = CStr(a(i + 1, cols(j)))
            Else
                fmt = "#,##0.00"
                If j = 2 Then fmt = "0.00%"
                out(i, j) = Format$(a(i + 1, cols(j)), fmt)
            End If
        Next j
    Next i
    lstResults.Clear: lstResults.list = out
End Sub
Private Sub btnResultsRefresh_Click()
    On Error GoTo Failed
    If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2))) <> "ACTUEL" Then
        CompareScenarios True
    End If
    Application.Calculate
    RefreshResults
    lblStatus.Caption = "Resultats actualises a partir des trois scenarios."
    Exit Sub
Failed:
    MsgBox "Actualisation des resultats : " & Err.Description, vbExclamation
End Sub
Private Sub btnSummary_Click()
    ShowSummaryReport
End Sub
Private Sub btnEvolution_Click()
    ShowEvolutionReport
End Sub
Private Sub btnGuideSheet_Click()
    Dim guide As Worksheet
    Set guide = ThisWorkbook.Worksheets("Guide")
    guide.Visible = xlSheetVisible
    Me.Hide
    guide.Activate
End Sub

Private Sub mReportSelector_Change()
    If mRefreshing Or mLayoutBusy Or mReportBusy Or Not mReady Then Exit Sub
    RenderReport
End Sub

Private Function ScenarioLabel(ByVal code As Variant) As String
    Dim row As Long
    row = FindRow("tblScenarios", CStr(code))
    If row > 0 Then ScenarioLabel = CStr(Tbl("tblScenarios").DataBodyRange.Cells(row, 2).Value2) Else ScenarioLabel = CStr(code)
End Function

Private Sub SetReportView(ByVal index As Long)
    mReportBusy = True: mReportSelector.ListIndex = index: mReportBusy = False
    If tabsMain.value = 10 Then RenderReport Else SelectPage 10
End Sub

Private Sub ShowSummaryReport(): SetReportView 0: End Sub
Private Sub ShowEvolutionReport(): SetReportView 1: End Sub
Private Sub ShowCustomReport(): SetReportView 5: End Sub

Public Sub RenderReport()
    On Error GoTo Failed
    Dim p As Object, a As Variant, data() As Variant, headers As Variant, widths As String
    Dim title As String, note As String, i As Long, j As Long, n As Long, cols As Variant
    Dim sums(1 To 3, 1 To 2, 1 To 4) As Double, counts(1 To 3, 1 To 2) As Long, code As Long, mode As Long, k As Long
    If mReportBusy Then Exit Sub
    mReportBusy = True
    Set p = tabsMain.Pages("pgRapports")
    Select Case mReportSelector.ListIndex
        Case 0
            title = "Synthese des trois scenarios"
            headers = Array("Scenario", "Cout ($)", "Service", "IC bas ($)", "IC haut ($)", "Stock moyen", "Ruptures (u)", "Commandes")
            widths = "130;100;90;100;100;100;100;100"
            a = Tbl("tblExport_Stats").DataBodyRange.Value2: n = UBound(a, 1)
            ReDim data(0 To n - 1, 0 To 7)
            cols = Array(1, 2, 6, 4, 5, 12, 10, 11)
            For i = 1 To n
                For j = 0 To 7
                    If j = 0 Then
                        data(i - 1, j) = ScenarioLabel(a(i, 1))
                    ElseIf j = 2 Then
                        data(i - 1, j) = Format$(a(i, cols(j)), "0.00%")
                    Else
                        data(i - 1, j) = Format$(a(i, cols(j)), "#,##0.00")
                    End If
                Next j
            Next i
            note = "Moyennes sur R repetitions. IC a 95 % du cout moyen. Ruptures en unites non servies. Stock de fin de journee."
        Case 1
            title = "Evolution quotidienne du stock"
            headers = Array("Scenario", "Jour", "Demande", "Stock final", "Ventes", "Ruptures (u)", "Commande (u)")
            widths = "140;65;100;100;100;110;110"
            a = Tbl("tblExport_PowerBI").DataBodyRange.Value2: n = UBound(a, 1)
            ReDim data(0 To n - 1, 0 To 6)
            For i = 1 To n
                data(i - 1, 0) = ScenarioLabel(a(i, 1))
                For j = 1 To 6: data(i - 1, j) = Format$(a(i, j + 1), "#,##0"): Next j
            Next i
            note = "Run 1, memes demandes pour les trois scenarios. Faites defiler verticalement pour consulter tous les jours."
        Case 2
            title = "Disponibilite et decomposition des couts"
            headers = Array("Scenario", "Disponible", "Jours rupture", "Max. (jours)", "Couverture (j)", "Possession ($)", "Commande ($)", "Rupture ($)")
            widths = "125;90;100;85;100;110;110;110"
            ThisWorkbook.Worksheets("Synthese").Calculate
            a = ThisWorkbook.Worksheets("Synthese").Range("B22:J24").Value2: n = 3
            ReDim data(0 To 2, 0 To 7): cols = Array(1, 2, 3, 4, 5, 7, 8, 9)
            For i = 1 To 3
                data(i - 1, 0) = a(i, 1): data(i - 1, 1) = Format$(a(i, 2), "0.00%")
                For j = 2 To 7
                    If IsEmpty(a(i, cols(j))) Or Len(CStr(a(i, cols(j)))) = 0 Then
                        data(i - 1, j) = "n.a."
                    Else
                        data(i - 1, j) = Format$(a(i, cols(j)), "#,##0.00")
                    End If
                Next j
            Next i
            note = "Disponibilite = part des jours sans rupture. Max. = plus longue sequence de rupture. Couverture = stock moyen / mu."
        Case 3
            title = "Comparaisons appariees a 95 %"
            headers = Array("Scenario A", "Scenario B", "Ecart A-B ($)", "IC bas ($)", "IC haut ($)", "Conclusion")
            widths = "120;120;120;110;110;220"
            a = Tbl("tblExport_Apparie").DataBodyRange.Value2: n = UBound(a, 1)
            ReDim data(0 To n - 1, 0 To 5)
            For i = 1 To n
                data(i - 1, 0) = ScenarioLabel(a(i, 1)): data(i - 1, 1) = ScenarioLabel(a(i, 2))
                data(i - 1, 2) = Format$(a(i, 3), "#,##0.00")
                data(i - 1, 3) = Format$(a(i, 5), "#,##0.00"): data(i - 1, 4) = Format$(a(i, 6), "#,##0.00")
                data(i - 1, 5) = a(i, 7)
            Next i
            note = "Difference de cout A - B sur les memes demandes. Si l'IC contient zero, aucune difference significative n'est mise en evidence."
        Case 4
            title = "Sensibilite aux commandes multiples"
            headers = Array("Scenario", "Multiples", "Cout moyen ($)", "Service", "Ruptures (u)", "Commandes")
            widths = "140;90;150;120;140;120"
            a = Tbl("tblExport_Sensibilite").DataBodyRange.Value2
            For i = 1 To UBound(a, 1)
                code = CLng(a(i, 1)): If a(i, 3) = "Oui" Then mode = 1 Else mode = 2
                counts(code, mode) = counts(code, mode) + 1
                sums(code, mode, 1) = sums(code, mode, 1) + a(i, 4)
                sums(code, mode, 2) = sums(code, mode, 2) + a(i, 5)
                sums(code, mode, 3) = sums(code, mode, 3) + a(i, 9)
                sums(code, mode, 4) = sums(code, mode, 4) + a(i, 7)
            Next i
            n = 6: ReDim data(0 To 5, 0 To 5): k = 0
            For code = 1 To 3: For mode = 1 To 2
                data(k, 0) = ScenarioLabel(code): data(k, 1) = IIf(mode = 1, "Oui", "Non")
                Require counts(code, mode) > 0, "Sensibilite incomplete. Relancez la comparaison."
                For j = 1 To 4
                    If j = 2 Then data(k, j + 1) = Format$(sums(code, mode, j) / counts(code, mode), "0.00%") Else data(k, j + 1) = Format$(sums(code, mode, j) / counts(code, mode), "#,##0.00")
                Next j
                k = k + 1
            Next mode: Next code
            note = "Moyennes des repetitions, avec la meme matrice de demandes en Oui et Non. Une commande au maximum par jour."
        Case 5
            title = "Simulation personnalisee s / Q"
            headers = Array("Repetition", "Cout ($)", "Service", "Ruptures (u)", "Commandes")
            widths = "100;150;130;150;130"
            With ThisWorkbook.Worksheets("Feuille_Controle")
                If CStr(.Range("L19").Value2) <> CustomSignature() Then
                    note = "Aucun resultat pour ces parametres. Dans Simulation, cliquez Simuler les valeurs saisies s / Q."
                    n = 0
                Else
                    n = CLng(Param("R")): a = .Range("F19").Resize(n, 5).Value2
                    ReDim data(0 To n - 1, 0 To 4)
                    For i = 1 To n
                        For j = 1 To 5
                            If j = 3 Then data(i - 1, j - 1) = Format$(a(i, j), "0.00%") Else data(i - 1, j - 1) = Format$(a(i, j), "#,##0.00")
                        Next j
                    Next i
                    note = "Resultats par repetition. s=" & Param("s") & ", Q=" & Param("Q") & ". Meme graine que la comparaison."
                End If
            End With
    End Select
    p.Controls("lblReportTitle").Caption = title
    p.Controls("lblReportSubtitle").Caption = "Produit " & Param("ProduitSimulation") & "   -   N=" & Param("N") & " jours   -   R=" & Param("R") & "   -   Objectif " & Format$(Param("SeuilService"), "0%")
    p.Controls("lblReportContext").Caption = ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2 & "   -   Decision : " & ThisWorkbook.Worksheets("Export_Decision").Range("A2").Value2
    If mReportSelector.ListIndex <> 5 Then
        p.Controls("lblReportSubtitle").Caption = "Dernier calcul : " & Format$(ThisWorkbook.Worksheets("Export_Decision").Range("C2").Value2, "dd/mm/yyyy hh:nn") & "   -   " & CStr(n) & " lignes"
        p.Controls("lblReportContext").Caption = p.Controls("lblReportContext").Caption & "   -   Produit " & Param("ProduitSimulation") & ", N=" & Param("N") & ", R=" & Param("R")
    End If
    If ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2 <> "ACTUEL" And mReportSelector.ListIndex <> 5 Then
        p.Controls("lblReportContext").Caption = "RESULTATS PERIMES. Relancez Comparer les scenarios avant d'interpreter ces valeurs."
        p.Controls("lblReportContext").ForeColor = RGB(170, 70, 35)
    Else
        p.Controls("lblReportContext").ForeColor = RGB(54, 98, 100)
    End If
    p.Controls("lblReportNotes").Caption = note
    mReportWidths = widths: mReportHeaders = headers
    FitReportColumns p, widths, headers
    p.Controls("lstReport").Clear
    If n > 0 Then p.Controls("lstReport").list = data
    mReportBusy = False
    Exit Sub
Failed:
    mReportBusy = False
    lblStatus.Caption = "Rapport : " & Err.Description
    MsgBox lblStatus.Caption, vbExclamation
End Sub

Public Sub ShowScenarioResults()
    If Not Me.Visible Then Me.Show vbModeless
    ShowSummaryReport
End Sub

Public Sub ShowStockEvolution()
    If Not Me.Visible Then Me.Show vbModeless
    ShowEvolutionReport
End Sub
Private Sub btnExport_Click()
    On Error GoTo Failed
    If Not ApplySettings() Then Exit Sub
    ExportAll True
    lblStatus.Caption = "CSV actualises dans le dossier exports."
    MsgBox lblStatus.Caption, vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub btnPowerBI_Click()
    If Not ApplySettings() Then Exit Sub
    OpenPowerBI
End Sub
Private Sub LoadGuide()
    Dim sh As Worksheet, a As Variant, lines As Variant, i As Long, j As Long, text As String, line As String
    Set sh = ThisWorkbook.Worksheets("Guide")
    lines = Array( _
        "1. Activer les macros du classeur pour utiliser le Menu, les formulaires et les simulations.", _
        "2. Gerer les produits, fournisseurs, commandes, stocks et receptions depuis le formulaire.", _
        "3. Choisir un produit dans Parametres recalcule sa demande, son delai, les trois suggestions s / Q et S0.", _
        "4. Suggestions modifiables. Recalibrer les regenere. Comparer utilise les memes demandes et le meme S0 pour les trois scenarios.", _
        "5. La comparaison porte sur le produit choisi. Apres un changement de produit, relancer la comparaison pour actualiser les resultats.", _
        "6. Pour Power BI, cliquer Analyse decisionnelle / Power BI puis Actualiser le rapport. Les chiffres affiches correspondent alors aux resultats exportes pour le produit choisi.", _
        "Stock moyen : moyenne des stocks de fin de journee. Ruptures : ventes perdues, sans reliquat.", _
        "Ordre quotidien : reception, vente, puis commande sur position stock physique + transit.", _
        "Une commande maximum par jour. Les commandes en transit sont autorisees par defaut.", _
        "Intervalles a 95 % : Student, R-1 degres de liberte. Comparaisons sur differences appariees.", _
        "Decision : service moyen >= seuil, puis cout moyen minimal. Un IC incluant zero indique une equivalence statistique.", _
        "Historique : quatre produits calibres sur leurs propres commandes DataCo non annulees, sur les jours observes.", _
        "Marge par unite moyenne, pertes incluses. Cpos=0,03 est une hypothese de gestion.", _
        "Le delai reste constant pendant chaque simulation. Sa valeur initiale est calibree pour le produit choisi.", _
        "Clients : identifiants fictifs de demonstration. L'historique source reste immuable via les formulaires.", _
        "Dates de saisie : jj/mm/aaaa. Quantites entieres positives. Reception unique par approvisionnement.", _
        "Nouveaux KPI : Synthese B21:J24 et ADDENDUM_KPI.md. Couverture vide si mu=0. Seuil de fiabilite = objectif de service en volume.")
    sh.Range("B2").Value2 = "Utilisation et methode"
    sh.Range("B5:B22").ClearContents
    For i = LBound(lines) To UBound(lines)
        sh.Cells(i + 5, 2).Value2 = lines(i)
    Next i
    a = sh.UsedRange.Value2
    For i = 1 To UBound(a, 1)
        line = ""
        For j = 1 To UBound(a, 2)
            If Not IsError(a(i, j)) Then
                If Len(CStr(a(i, j))) > 0 Then
                    If Len(line) > 0 Then line = line & "  "
                    line = line & CStr(a(i, j))
                End If
            End If
        Next j
        If Len(line) > 0 Then text = text & line & vbCrLf & vbCrLf
    Next i
    txtGuide.value = text
End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set mDesignButtons = Nothing
    If CloseMode = 0 Or CloseMode = 1 Then ShowExcelHome
End Sub
Private Sub btnHome_Click()
    SelectPage 0
End Sub

Private Sub nav1_Click()
SelectPage 1
End Sub

Private Sub nav2_Click()
SelectPage 2
End Sub

Private Sub nav3_Click()
SelectPage 3
End Sub

Private Sub nav4_Click()
SelectPage 4
End Sub

Private Sub nav5_Click()
SelectPage 5
End Sub

Private Sub nav6_Click()
SelectPage 6
End Sub

Private Sub nav7_Click()
SelectPage 7
End Sub

Private Sub nav8_Click()
SelectPage 8
End Sub

Private Sub nav9_Click()
SelectPage 9
End Sub

Private Sub pr_Initialize()
pr_RefreshList
End Sub

Private Sub pr_btnSearch_Click()
pr_RefreshList
End Sub

Private Sub pr_btnNew_Click()
pr_txtKey.Locked = False: pr_txtKey.value = "": pr_txtName.value = "": pr_txtPrice.value = ""
End Sub

Private Sub pr_lstRecords_Click()
On Error GoTo Failed
If pr_lstRecords.ListIndex < 0 Then Exit Sub
Dim i As Long
i = FindRow("tblT_Produits", CStr(pr_lstRecords.list(pr_lstRecords.ListIndex, 0)))
With Tbl("tblT_Produits").DataBodyRange
pr_txtKey.value = .Cells(i, 1).Value2: pr_txtName.value = .Cells(i, 2).Value2: pr_txtPrice.value = .Cells(i, 3).Value2
End With
pr_txtKey.Locked = True
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub pr_btnSave_Click()
On Error GoTo Failed
SaveProduct CStr(pr_txtKey.value), CStr(pr_txtName.value), NumberValue(pr_txtPrice.value, "Prix")
pr_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub pr_btnDelete_Click()
On Error GoTo Failed
DeleteProduct CStr(pr_txtKey.value)
pr_txtKey.value = "": pr_txtName.value = "": pr_txtPrice.value = "": pr_txtKey.Locked = False
pr_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub pr_RefreshList()
FillRecords pr_lstRecords, "tblT_Produits", Array(1, 2, 3), CStr(pr_txtSearch.value)
End Sub
Private Sub fu_Initialize()
fu_RefreshList
End Sub

Private Sub fu_btnSearch_Click()
fu_RefreshList
End Sub

Private Sub fu_btnNew_Click()
fu_txtKey.Locked = False: fu_txtKey.value = "": fu_txtName.value = "": fu_txtDelay.value = "": fu_txtCost.value = ""
End Sub

Private Sub fu_lstRecords_Click()
On Error GoTo Failed
If fu_lstRecords.ListIndex < 0 Then Exit Sub
Dim i As Long
i = FindRow("tblT_Fournisseurs", CStr(fu_lstRecords.list(fu_lstRecords.ListIndex, 0)))
With Tbl("tblT_Fournisseurs").DataBodyRange
fu_txtKey.value = .Cells(i, 1).Value2: fu_txtName.value = .Cells(i, 2).Value2
fu_txtDelay.value = .Cells(i, 3).Value2: fu_txtCost.value = .Cells(i, 4).Value2
End With
fu_txtKey.Locked = True
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub fu_btnSave_Click()
On Error GoTo Failed
SaveSupplier CStr(fu_txtKey.value), CStr(fu_txtName.value), NumberValue(fu_txtDelay.value, "Delai"), NumberValue(fu_txtCost.value, "Cout")
fu_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub fu_btnDelete_Click()
On Error GoTo Failed
DeleteSupplier CStr(fu_txtKey.value)
fu_txtKey.Locked = False: fu_txtKey.value = "": fu_txtName.value = ""
fu_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub fu_RefreshList()
FillRecords fu_lstRecords, "tblT_Fournisseurs", Array(1, 2, 3, 4), CStr(fu_txtSearch.value)
End Sub
Private Sub co_Initialize()
FillProducts co_cboProduct
co_txtDate.value = Format$(Date, "dd/mm/yyyy")
co_txtClient.value = "Client 001"
co_txtKey.Locked = True
co_RefreshList
End Sub

Private Sub co_btnSearch_Click()
co_RefreshList
End Sub

Private Sub co_btnNew_Click()
co_txtKey.value = "": co_txtQty.value = "": co_txtDate.value = Format$(Date, "dd/mm/yyyy")
End Sub

Private Sub co_lstRecords_Click()
On Error GoTo Failed
If co_lstRecords.ListIndex < 0 Then Exit Sub
co_txtKey.value = co_lstRecords.list(co_lstRecords.ListIndex, 0)
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub co_btnSave_Click()
On Error GoTo Failed
Dim key As String
key = CreateOrder(CStr(co_cboProduct.value), CLng(NumberValue(co_txtQty.value, "Quantite", 1, True)), ParseDate(CStr(co_txtDate.value)), CStr(co_txtClient.value))
co_txtKey.value = key
MsgBox "Commande " & key & " enregistree. Les unites indisponibles sont des ventes perdues.", vbInformation
co_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub co_btnCancel_Click()
On Error GoTo Failed
CancelOrder CStr(co_txtKey.value)
MsgBox "Commande annulee. Stock vendu restitue.", vbInformation
co_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub co_RefreshList()
FillRecords co_lstRecords, "tblT_Commandes", Array(1, 2, 4, 5, 8), CStr(co_txtSearch.value)
End Sub
Private Sub ap_Initialize()
FillProducts ap_cboProduct
FillSuppliers ap_cboSupplier
ap_txtDate.value = Format$(Date, "dd/mm/yyyy")
ap_txtReceipt.value = Format$(Date, "dd/mm/yyyy")
ap_txtKey.Locked = True
ap_RefreshList
End Sub

Private Sub ap_btnSearch_Click()
ap_RefreshList
End Sub

Private Sub ap_btnNew_Click()
ap_txtKey.value = "": ap_txtQty.value = ""
End Sub

Private Sub ap_lstRecords_Click()
On Error GoTo Failed
If ap_lstRecords.ListIndex < 0 Then Exit Sub
ap_txtKey.value = ap_lstRecords.list(ap_lstRecords.ListIndex, 0)
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub ap_btnSave_Click()
On Error GoTo Failed
ap_txtKey.value = CreateSupply(CStr(ap_cboSupplier.value), CStr(ap_cboProduct.value), CLng(NumberValue(ap_txtQty.value, "Quantite", 1, True)), ParseDate(CStr(ap_txtDate.value)))
MsgBox "Approvisionnement cree.", vbInformation
ap_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub ap_btnReceive_Click()
On Error GoTo Failed
ReceiveSupply CStr(ap_txtKey.value), ParseDate(CStr(ap_txtReceipt.value))
MsgBox "Reception enregistree et stock actualise.", vbInformation
ap_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub ap_RefreshList()
FillRecords ap_lstRecords, "tblT_Approvisionnements", Array(1, 3, 4, 5, 8), CStr(ap_txtSearch.value)
End Sub
Private Sub st_Initialize()
FillProducts st_cboProduct
st_txtDelta.value = "0"
st_txtThreshold.value = "260"
st_RefreshList
End Sub

Private Sub st_btnSearch_Click()
st_RefreshList
End Sub

Private Sub st_btnNew_Click()
st_txtDelta.value = "0"
End Sub

Private Sub st_lstRecords_Click()
On Error GoTo Failed
If st_lstRecords.ListIndex < 0 Then Exit Sub
st_cboProduct.value = st_lstRecords.list(st_lstRecords.ListIndex, 0)
st_txtThreshold.value = st_lstRecords.list(st_lstRecords.ListIndex, 2)
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub st_btnSave_Click()
On Error GoTo Failed
Require IsNumeric(st_txtDelta.value), "Mouvement numerique obligatoire."
Require CDbl(st_txtDelta.value) = Fix(CDbl(st_txtDelta.value)), "Mouvement entier obligatoire."
AdjustStock CStr(st_cboProduct.value), CLng(st_txtDelta.value), CLng(NumberValue(st_txtThreshold.value, "Seuil", 0, True))
st_txtDelta.value = "0"
st_RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub st_btnTable_Click()
st_RefreshList
SelectPage 5
lblStatus.Caption = "Le tableau des stocks est affiche dans l'onglet Stocks."
End Sub

Private Sub st_RefreshList()
FillRecords st_lstRecords, "tblT_Stock", Array(1, 2, 3, 4), CStr(st_txtSearch.value)
End Sub
Private Sub st_cboProduct_Change()
If mRefreshing Then Exit Sub
Dim i As Long
i = FindRow("tblT_Stock", CStr(st_cboProduct.value))
If i > 0 Then st_txtThreshold.value = Tbl("tblT_Stock").DataBodyRange.Cells(i, 3).Value2
End Sub








Private Sub InstallDesignControls()
    Dim i As Long, c As MSForms.CommandButton, handler As CDataCoDesignButton
    Dim names As Variant, actions As Variant, captions As Variant
    Set mDesignButtons = New Collection
    For i = 0 To tabsMain.Pages.count - 1
        mPageZoom(i) = 80
        Set c = Me.Controls.Add("Forms.CommandButton.1", "designNav" & i, True)
        c.Caption = tabsMain.Pages(i).Caption: c.Tag = "page:" & i
        Set handler = New CDataCoDesignButton
        Set handler.Button = c: Set handler.Owner = Me
        mDesignButtons.Add handler
    Next i
    mPageZoom(0) = 100
    names = Array("designZoomOut", "designZoomReset", "designZoomIn", "designZoomFit")
    actions = Array("zoom:-10", "zoom:100", "zoom:10", "fit")
    captions = Array("-", "100%", "+", "Ajuster")
    For i = 0 To 3
        Set c = Me.Controls.Add("Forms.CommandButton.1", CStr(names(i)), True)
        c.Caption = captions(i): c.Tag = actions(i)
        Set handler = New CDataCoDesignButton
        Set handler.Button = c: Set handler.Owner = Me
        mDesignButtons.Add handler
    Next i
End Sub

Public Sub HandleDesignAction(ByVal action As String)
    If Left$(action, 5) = "page:" Then
        SelectPage CLng(Mid$(action, 6))
    Else
        Select Case action
            Case "fit": mContentZoom = FitDataCoZoom(Me)
            Case "zoom:100": mContentZoom = 100
            Case "zoom:-10": mContentZoom = Application.Max(50, mContentZoom - 10)
            Case "zoom:10": mContentZoom = Application.Min(150, mContentZoom + 10)
        End Select
        mLayoutBusy = True
        mPageZoom(tabsMain.value) = mContentZoom
        ZoomDataCoPages Me, mContentZoom
        mLayoutBusy = False
    End If
End Sub

Private Sub ApplyDesignZoom()
    Dim i As Long
    For i = 0 To 10
        ZoomDataCoPages Me, mPageZoom(i), i
    Next i
    mContentZoom = mPageZoom(tabsMain.value)
    Me.Controls("designZoomReset").Caption = CStr(mContentZoom) & "%"
End Sub

