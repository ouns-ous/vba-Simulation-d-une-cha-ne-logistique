Option Explicit

Public Sub OpenProducts(): OpenApplication 1: End Sub
Public Sub OpenSuppliers(): OpenApplication 2: End Sub
Public Sub OpenOrders(): OpenApplication 3: End Sub
Public Sub OpenSupply(): OpenApplication 4: End Sub
Public Sub OpenStock(): OpenApplication 5: End Sub
Public Sub OpenControl(): OpenApplication 6: End Sub
Public Sub OpenSummary(): OpenApplication 8: End Sub
Public Sub OpenEvolution()
    OpenApplication 8
    F_Main.ShowStockEvolution
End Sub
Public Sub OpenGuide(): OpenApplication 9: End Sub
Public Sub GoMenu(): OpenApplication 0: End Sub

Public Sub UpdateOperationalKPI()
    Dim a As Variant, i As Long, alerts As Long, pending As Long, status As String
    If Not Tbl("tblT_Stock").DataBodyRange Is Nothing Then
        a = Tbl("tblT_Stock").DataBodyRange.Value2
        For i = 1 To UBound(a, 1)
            If CDbl(a(i, 2)) < CDbl(a(i, 3)) Then alerts = alerts + 1
        Next i
    End If
    If Not Tbl("tblT_Approvisionnements").DataBodyRange Is Nothing Then
        a = Tbl("tblT_Approvisionnements").DataBodyRange.Value2
        For i = 1 To UBound(a, 1)
            status = LCase$(Replace$(Trim$(CStr(a(i, 8))), " ", ""))
            If status = "entransit" Or status = "enattente" Then pending = pending + 1
        Next i
    End If
    With ThisWorkbook.Worksheets("Menu")
        .Range("B35").Value2 = "Produits sous seuil": .Range("E35").Value2 = alerts
        .Range("B36").Value2 = "Approvisionnements en attente": .Range("E36").Value2 = pending
    End With
End Sub

Public Sub ProtectData()
    Dim names As Variant, v As Variant
    names = Array("T_Produits", "T_Fournisseurs", "T_Commandes", "T_Stock", "T_Approvisionnements", "T_Mouvements")
    For Each v In names
        ThisWorkbook.Worksheets(CStr(v)).Protect UserInterfaceOnly:=True, AllowFiltering:=True
    Next v
End Sub

Public Sub FillProducts(ByVal combo As Object)
    Dim a As Variant, i As Long
    combo.Clear: combo.ColumnCount = 2: combo.BoundColumn = 1: combo.TextColumn = 2
    combo.ColumnWidths = "45;210"
    a = Tbl("tblT_Produits").DataBodyRange.Value2
    For i = 1 To UBound(a, 1)
        combo.AddItem a(i, 1): combo.list(combo.ListCount - 1, 1) = a(i, 2)
    Next i
    If combo.ListCount > 0 Then combo.ListIndex = 0
End Sub

Public Sub FillSuppliers(ByVal combo As Object)
    Dim a As Variant, i As Long
    combo.Clear: combo.ColumnCount = 2: combo.BoundColumn = 1: combo.TextColumn = 2
    combo.ColumnWidths = "45;210"
    a = Tbl("tblT_Fournisseurs").DataBodyRange.Value2
    For i = 1 To UBound(a, 1)
        combo.AddItem a(i, 1): combo.list(combo.ListCount - 1, 1) = a(i, 2)
    Next i
    If combo.ListCount > 0 Then combo.ListIndex = 0
End Sub

Public Sub FillRecords(ByVal list As Object, ByVal tableName As String, ByVal columns As Variant, Optional ByVal search As String = "", Optional ByVal maximum As Long = 1000)
    Dim a As Variant, i As Long, j As Long, count As Long, text As String, match As Boolean
    list.Clear: list.ColumnCount = UBound(columns) - LBound(columns) + 1
    If Tbl(tableName).DataBodyRange Is Nothing Then Exit Sub
    a = Tbl(tableName).DataBodyRange.Value2
    For i = UBound(a, 1) To 1 Step -1
        If CStr(a(i, 1)) <> "" Then
            match = search = ""
            If Not match Then
                For j = 1 To UBound(a, 2)
                    If InStr(1, CStr(a(i, j)), search, vbTextCompare) > 0 Then match = True: Exit For
                Next j
            End If
            If match Then
                list.AddItem CStr(a(i, columns(LBound(columns))))
                For j = LBound(columns) + 1 To UBound(columns)
                    text = CStr(a(i, columns(j)))
                    If columns(j) = 2 And (tableName = "tblT_Commandes" Or tableName = "tblT_Approvisionnements") Then text = Format$(CDate(a(i, 2)), "dd/mm/yyyy")
                    list.list(list.ListCount - 1, j - LBound(columns)) = text
                Next j
                count = count + 1: If count >= maximum Then Exit For
            End If
        End If
    Next i
End Sub

Public Sub ResetParameters()
    On Error GoTo Failed
    Dim selectedProduct As String
    selectedProduct = CStr(Param("ProduitSimulation"))
    If FindRow("tblT_Produits", selectedProduct) = 0 Then selectedProduct = "P001"
    SetParam "L", 3
    SetParam "Cpos", 0.03: SetParam "Ccmd", 12: SetParam "N", 365: SetParam "R", 30
    SetParam "Graine", 20261023: SetParam "SeuilService", 0.95: SetParam "CommandesMultiples", "Oui"
    SetParam "ProduitSimulation", selectedProduct: Calibrate True
    Call MarkStale: MsgBox "Parametres reinitialises et demande recalibree.", vbInformation: Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Function CsvValue(ByVal value As Variant, ByVal header As String) As String
    Dim text As String
    If IsEmpty(value) Then
        text = ""
    ElseIf (header = "Date" Or header = "DateCommande" Or header = "DatePrevue" Or header = "DateReception" Or header = "DateCalcul" Or header = "DateSimulation") And IsNumeric(value) Then
        text = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
        text = Replace$(CStr(value), Application.International(xlDecimalSeparator), ".")
    Else
        text = CStr(value)
    End If
    CsvValue = """" & Replace$(text, """", """""") & """"
End Function

Public Sub ExportOne(ByVal tableName As String, ByVal target As String)
    Dim lo As ListObject, a As Variant, h As Variant, lines() As String, i As Long, j As Long, row As String, stream As Object
    Set lo = Tbl(tableName): h = lo.HeaderRowRange.Value2
    ReDim lines(0 To lo.ListRows.count)
    For j = 1 To lo.ListColumns.count
        If j > 1 Then lines(0) = lines(0) & ","
        lines(0) = lines(0) & CsvValue(h(1, j), "")
    Next j
    If Not lo.DataBodyRange Is Nothing Then
        a = lo.DataBodyRange.Value2
        For i = 1 To UBound(a, 1)
            row = ""
            For j = 1 To UBound(a, 2)
                If j > 1 Then row = row & ","
                row = row & CsvValue(a(i, j), CStr(h(1, j)))
            Next j
            lines(i) = row
        Next i
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText Join(lines, vbCrLf): stream.SaveToFile target, 2: stream.Close
End Sub

Public Sub ExportAll(Optional ByVal silent As Boolean = False)
    On Error GoTo Failed
    Dim folder As String, tables As Variant, v As Variant
    If ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2 <> "ACTUEL" Then CompareScenarios True
    folder = ThisWorkbook.path & Application.PathSeparator & "exports"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    tables = Array("Export_PowerBI", "Export_KPI", "Export_Sensibilite", "Export_Stats", "Export_Apparie", "Export_Decision", "T_Produits", "T_Fournisseurs", "T_Commandes", "T_Stock", "T_Approvisionnements", "Scenarios")
    For Each v In tables
        If v = "Scenarios" Then
            ExportOne "tblScenarios", folder & Application.PathSeparator & v & ".csv"
        Else
            ExportOne "tbl" & v, folder & Application.PathSeparator & v & ".csv"
        End If
    Next v
    WriteSimulationContext folder
    UpdateOperationalKPI
    If Not silent Then MsgBox "Exports actualises dans " & folder, vbInformation
    Exit Sub
Failed:
    If silent Then Err.Raise Err.Number, Err.source, Err.Description Else MsgBox Err.Description, vbExclamation
End Sub

Private Sub WriteSimulationContext(ByVal folder As String)
    Dim product As String, row As Long, products As Variant, csv As String, stream As Object
    product = CStr(Param("ProduitSimulation"))
    row = FindRow("tblT_Produits", product)
    Require row > 0, "Produit de simulation absent du catalogue : " & product
    products = Tbl("tblT_Produits").DataBodyRange.Value2
    csv = CsvValue("CodeProduit", "") & "," & CsvValue("Designation", "") & "," & CsvValue("DateSimulation", "") & vbCrLf
    csv = csv & CsvValue(product, "CodeProduit") & "," & _
        CsvValue(products(row, 2), "Designation") & "," & _
        CsvValue(ThisWorkbook.Worksheets("Export_Decision").Range("C2").Value2, "DateSimulation")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText csv
    stream.SaveToFile folder & Application.PathSeparator & "Simulation_Contexte.csv", 2
    stream.Close
End Sub

Public Sub EnsureWorkbookVisible()
    Dim win As Window
    Application.Visible = True
    For Each win In ThisWorkbook.Windows
        win.Visible = True
        win.DisplayWorkbookTabs = True
        win.DisplayHorizontalScrollBar = False
        win.DisplayVerticalScrollBar = True
    Next win
    ThisWorkbook.Activate
End Sub

Public Sub ShowExcelHome()
    EnsureWorkbookVisible
    ThisWorkbook.Worksheets("Menu").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("Menu").Activate
End Sub

Public Sub ConfigurePowerBISource()
    Dim path As String, stream As Object, binary As Object, source As String, lines As Variant, i As Long, found As Boolean
    path = ThisWorkbook.path & "\PowerBI\DataCo.SemanticModel\definition\expressions.tmdl"
    Require Dir(path) <> "", "Parametre Power BI introuvable : " & path
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path: source = stream.ReadText: stream.Close
    lines = Split(Replace$(source, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If Left$(Trim$(lines(i)), 24) = "expression FolderExports" Then
            lines(i) = "expression FolderExports = """ & Replace$(ThisWorkbook.path, "\", "/") & _
                "/exports"" meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
            found = True
        End If
    Next i
    Require found, "Le parametre FolderExports est absent du modele Power BI."
    stream.Open: stream.WriteText Join(lines, vbCrLf)
    ' TMDL requires UTF-8 without BOM; ADODB text streams add a three-byte BOM.
    stream.Position = 0: stream.Type = 1: stream.Position = 3
    Set binary = CreateObject("ADODB.Stream")
    binary.Type = 1: binary.Open
    stream.CopyTo binary
    binary.SaveToFile path, 2: binary.Close: stream.Close
End Sub

Public Function PowerBIProjectPath() As String
    PowerBIProjectPath = ThisWorkbook.path & "\PowerBI\DataCo.pbip"
    Require Dir(PowerBIProjectPath) <> "", "Projet Power BI introuvable : " & PowerBIProjectPath
End Function

Public Sub OpenPowerBI()
    On Error GoTo Failed
    Dim path As String, exe As String, phase As String
    EnsureWorkbookVisible
    phase = "Recalcul de la simulation pour le produit selectionne": CompareScenarios True
    phase = "Export des donnees": ExportAll True
    phase = "Configuration de la source": ConfigurePowerBISource
    path = PowerBIProjectPath()
    exe = Environ$("ProgramW6432") & "\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
    If Dir(exe) = "" Then exe = Environ$("ProgramFiles") & "\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
    phase = "Ouverture du rapport"
    If Dir(exe) <> "" Then
        Shell """" & exe & """ """ & path & """", vbNormalFocus
    Else
        ThisWorkbook.FollowHyperlink path
    End If
    Exit Sub
Failed:
    MsgBox phase & " : " & Err.Description, vbExclamation
End Sub

Public Function SelfTest() As String
    On Error GoTo Failed
    Dim p As String, s As String, key As String, supply As String, i As Long, before As Long, caught As Boolean
    SaveProduct "TEST", "Produit test", 1: SaveProduct "TEST", "Produit test modifie", 2
    Require Tbl("tblT_Produits").DataBodyRange.Cells(FindRow("tblT_Produits", "TEST"), 3).Value2 = 2, "Modification produit echouee."
    SaveSupplier "TESTF", "Fournisseur test", 2, 3
    SaveSupplier "TESTF", "Fournisseur test modifie", 1, 4
    AdjustStock "TEST", 10, 5
    key = CreateOrder("TEST", 15, DateSerial(2026, 10, 10), "Client Test")
    i = FindRow("tblT_Commandes", key)
    Require Tbl("tblT_Commandes").DataBodyRange.Cells(i, 6).Value2 = 10, "Ventes incorrectes."
    Require Tbl("tblT_Commandes").DataBodyRange.Cells(i, 7).Value2 = 5, "Rupture incorrecte."
    CancelOrder key
    Require Tbl("tblT_Stock").DataBodyRange.Cells(FindRow("tblT_Stock", "TEST"), 2).Value2 = 10, "Annulation ne restaure pas le stock."
    On Error Resume Next
    CancelOrder key: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Require caught, "Double annulation non bloquee."
    supply = CreateSupply("TESTF", "TEST", 20, DateSerial(2026, 10, 10))
    ReceiveSupply supply, DateSerial(2026, 10, 11)
    Require Tbl("tblT_Stock").DataBodyRange.Cells(FindRow("tblT_Stock", "TEST"), 2).Value2 = 30, "Reception incorrecte."
    On Error Resume Next
    ReceiveSupply supply, DateSerial(2026, 10, 12): caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Require caught, "Double reception non bloquee."
    On Error Resume Next
    DeleteProduct "TEST": caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Require caught, "Suppression referencee non bloquee."
    Dim d() As Long, journal() As Variant, result As SimResult
    ReDim d(1 To 1, 1 To 4): ReDim journal(1 To 4, 1 To 7)
    d(1, 1) = 6: d(1, 2) = 6: d(1, 3) = 6: d(1, 4) = 6
    result = Simulate(d, 1, 4, 10, 4, 8, 2, True, 1, 2, 3, journal, True)
    Require journal(1, 4) = 4 And journal(2, 6) = 2 And journal(3, 4) = 2 And journal(4, 6) = 4, "Chronologie simulation incorrecte."
    Require result.lost = 6 And result.Orders = 2 And result.Cost = 28, "KPI deterministes incorrects."
    GenerateDemand d, 2, 10, 73, 17, 20261023
    before = d(1, 1): GenerateDemand d, 2, 10, 73, 17, 20261023
    Require d(1, 1) = before, "Graine non reproductible."
    GenerateDemand d, 2, 10, 0, 0, 1
    result = Simulate(d, 1, 10, 0, 0, 10, 1, True, 0, 0, 0, journal, False)
    Require result.Service = 1 And result.lost = 0 And result.Cost = 0, "Demande nulle incorrecte."
    Load F_Produit
    Require F_Produit.lstRecords.ListCount = 5, "Formulaire produits incomplet."
    Unload F_Produit
    Load F_Fournisseur
    Require F_Fournisseur.lstRecords.ListCount = 4, "Formulaire fournisseurs incomplet."
    Unload F_Fournisseur
    Load F_Commande
    Require F_Commande.lstRecords.ListCount = 1000, "Formulaire commandes incomplet."
    Unload F_Commande
    Load F_Appro
    Require F_Appro.lstRecords.ListCount = 1, "Formulaire approvisionnements incomplet."
    Unload F_Appro
    Load F_Stock
    Require F_Stock.lstRecords.ListCount = 5, "Formulaire stock incomplet."
    Unload F_Stock
    SaveProduct "DELETE", "Produit supprimable", 1
    DeleteProduct "DELETE"
    Require FindRow("tblT_Produits", "DELETE") = 0, "Suppression produit echouee."
    SaveSupplier "DELETEF", "Fournisseur supprimable", 1, 1
    DeleteSupplier "DELETEF"
    Require FindRow("tblT_Fournisseurs", "DELETEF") = 0, "Suppression fournisseur echouee."
    On Error Resume Next
    Dim invalidDate As Date: invalidDate = ParseDate("31/02/2026"): caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Require caught, "Date invalide acceptee."
    On Error Resume Next
    AdjustStock "TEST", -31, 5: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Require caught, "Stock negatif accepte."
    SaveSupplier "F001", "Global Sport Supply", 6, 15
    Calibrate True
    Require CLng(Param("L")) = 4 And CDbl(Param("Ccmd")) = 15, "Fournisseurs non lies a la calibration."
    SaveSupplier "F001", "Global Sport Supply", 4, 12
    Calibrate True
    Require CLng(Param("L")) = 3 And CDbl(Param("Ccmd")) = 12, "Calibration par defaut incorrecte."
    SetParam "R", 1
    On Error Resume Next
    Call ValidateParameters
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    SetParam "R", 30
    Require caught, "R=1 accepte pour les intervalles."
    ' The test workbook is discarded by the build script; demonstration data stay clean.
    SelfTest = "PASS : CRUD, ventes partielles, annulation, reception unique, references, simulation deterministe, graine, demande nulle, cinq UserForms, dates invalides."
    Exit Function
Failed:
    SelfTest = "FAIL " & Err.Number & " : " & Err.Description
End Function

Public Sub OpenMain(): OpenApplication 0: End Sub
Public Sub OpenSimulation(): OpenApplication 7: End Sub
Public Sub OpenExports(): OpenApplication 9: End Sub
Public Sub OpenApplication(Optional ByVal page As Long = 0)
    On Error GoTo Failed
    EnsureWorkbookVisible
    Application.WindowState = xlMaximized
    Load F_Main
    F_Main.SelectPage page
    If Not F_Main.Visible Then F_Main.Show vbModeless
    DoEvents
    F_Main.RefreshLayout
    F_Main.RefreshPage
    Exit Sub
Failed:
    MsgBox "Interface DataCo : " & Err.Description, vbExclamation
End Sub



