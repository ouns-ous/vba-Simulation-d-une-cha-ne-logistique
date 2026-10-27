Attribute VB_Name = "modFormLayout"
Option Explicit
Private mPageWidth As Single
Private mPageHeight As Single

Private Function Ink() As Long
    Ink = RGB(39, 49, 63)
End Function
Private Function Muted() As Long
    Muted = RGB(110, 121, 139)
End Function
Private Function Blue() As Long
    Blue = RGB(85, 101, 160)
End Function
Private Function Canvas() As Long
    Canvas = RGB(243, 245, 250)
End Function
Private Function BrandNavy() As Long
    BrandNavy = RGB(255, 255, 255)
End Function
Private Function Accent() As Long
    Accent = RGB(85, 101, 160)
End Function

Private Sub Place(ByVal host As Object, ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Bounds host.Controls(name), x, y, w, h
End Sub

Private Sub Backdrop(ByVal host As Object, ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal color As Long)
    Dim c As Object
    On Error Resume Next
    Set c = host.Controls(name)
    On Error GoTo 0
    If c Is Nothing Then Set c = host.Controls.Add("Forms.Label.1", name, True)
    c.Caption = ""
    c.BackStyle = 1
    c.BackColor = color
    c.BorderStyle = 0
    Bounds c, x, y, w, h
    c.ZOrder 1
End Sub

Private Sub TextStyle(ByVal c As Object, Optional ByVal size As Single = 11, Optional ByVal bold As Boolean = False)
    c.Font.name = "Segoe UI"
    c.Font.size = size
    c.Font.bold = bold
    c.ForeColor = Ink()
End Sub

Private Sub ButtonStyle(ByVal c As Object, Optional ByVal primary As Boolean = False, Optional ByVal danger As Boolean = False)
    TextStyle c, 11, True
    c.Visible = True
    c.TakeFocusOnClick = False
    If primary Then
        c.BackColor = Blue(): c.ForeColor = RGB(255, 255, 255)
    ElseIf danger Then
        c.BackColor = RGB(247, 248, 250): c.ForeColor = RGB(148, 69, 43)
    Else
        c.BackColor = RGB(247, 248, 250): c.ForeColor = Ink()
    End If
End Sub

Private Function Heading(ByVal host As Object, ByVal name As String) As Object
    Dim c As Object
    On Error Resume Next
    Set c = host.Controls(name)
    On Error GoTo 0
    If c Is Nothing Then Set c = host.Controls.Add("Forms.Label.1", name, True)
    c.BackStyle = 0
    c.WordWrap = True
    Set Heading = c
End Function

Private Sub PageTitle(ByVal p As Object, ByVal name As String, ByVal w As Single)
    Place p, name, 16, 8, w - 32, 28
    TextStyle p.Controls(name), 17, True
    p.Controls(name).ForeColor = Ink()
End Sub

Private Sub StylePage(ByVal p As Object)
    Dim c As Object
    For Each c In p.Controls
        Select Case TypeName(c)
            Case "Label"
                TextStyle c, 10.5, True
                c.BackStyle = 0
                c.WordWrap = True
            Case "TextBox", "ComboBox"
                TextStyle c, 11, True
                c.BackColor = RGB(255, 255, 255)
                c.SpecialEffect = 0
                c.BorderStyle = 1
                c.BorderColor = RGB(205, 211, 222)
                If TypeName(c) = "TextBox" Then
                    If c.Locked Then c.BackColor = RGB(255, 255, 255)
                End If
            Case "ListBox"
                TextStyle c, 10.5
                c.BackColor = RGB(255, 255, 255)
                c.BorderStyle = 1
                c.BorderColor = RGB(205, 214, 228)
                c.SpecialEffect = 0
                c.IntegralHeight = False
            Case "CommandButton"
                ButtonStyle c
        End Select
    Next c
End Sub

Private Sub CatalogPage(ByVal p As Object, ByVal prefix As String, ByVal fields As Variant, ByVal headers As Variant)
    Dim w As Single, h As Single, leftW As Single, rx As Single, rightW As Single
    Dim i As Long, c As Object, name As String, colW As Single, x As Single, widths As String
    Dim fieldX As Single, fieldW As Single, fieldRow As Long
    w = mPageWidth: h = mPageHeight
    leftW = Application.Max(268, Application.Min(332, w * 0.33))
    rx = leftW + 44: rightW = w - rx - 18
    PageTitle p, prefix & "lblTitle", w
    For i = LBound(fields) To UBound(fields)
        name = CStr(fields(i))
        fieldRow = i: fieldX = 18: fieldW = leftW
        If prefix = "ap_" Then
            If name = "txtQty" Or name = "txtDate" Then
                fieldRow = 3: fieldW = (leftW - 10) / 2
                If name = "txtDate" Then fieldX = 28 + fieldW
            End If
            If name = "txtReceipt" Then fieldRow = 4
        End If
        Place p, prefix & "lbl" & name, fieldX, 43 + fieldRow * 48, fieldW, 16
        Place p, prefix & name, fieldX, 61 + fieldRow * 48, fieldW, 28
    Next i
    Place p, prefix & "lblSearch", rx, 43, rightW, 16
    p.Controls(prefix & "lblSearch").Caption = "Rechercher dans la liste"
    Place p, prefix & "txtSearch", rx, 61, rightW - 132, 28
    Place p, prefix & "btnSearch", rx + rightW - 122, 61, 122, 28
    Place p, prefix & "lstRecords", rx, 120, rightW, h - 174
    Place p, prefix & "lblNote", rx, h - 44, rightW, 36
    p.Controls(prefix & "lblNote").ForeColor = Muted()
    p.Controls(prefix & "lblColumns").Visible = False
    widths = "": x = rx
    For i = LBound(headers) To UBound(headers)
        Select Case prefix
            Case "pr_"
                If i = 1 Then colW = rightW - 174 Else colW = 80
            Case "fu_"
                Select Case i
                    Case 0: colW = 70
                    Case 1: colW = rightW - 254
                    Case Else: colW = 85
                End Select
            Case "co_"
                Select Case i
                    Case 0: colW = 102
                    Case 1: colW = 82
                    Case 2: colW = 76
                    Case 3: colW = 70
                    Case Else: colW = rightW - 344
                End Select
            Case "ap_"
                Select Case i
                    Case 0: colW = 102
                    Case 1: colW = 92
                    Case 2: colW = 76
                    Case 3: colW = 70
                    Case Else: colW = rightW - 354
                End Select
            Case Else
                colW = (rightW - 14) / (UBound(headers) + 1)
        End Select
        Set c = Heading(p, prefix & "head" & i)
        c.Caption = CStr(headers(i))
        Bounds c, x + 3, 99, colW - 3, 19
        TextStyle c, 10, True
        x = x + colW
        If Len(widths) > 0 Then widths = widths & ";"
        widths = widths & CStr(Int(colW))
    Next i
    p.Controls(prefix & "lstRecords").ColumnWidths = widths
    Place p, prefix & "btnSave", 18, h - 76, (leftW - 10) / 2, 34
    ButtonStyle p.Controls(prefix & "btnSave"), True
    name = ""
    Select Case prefix
        Case "pr_", "fu_": name = "btnDelete"
        Case "co_": name = "btnCancel"
        Case "ap_": name = "btnReceive"
        Case "st_": name = "btnTable"
    End Select
    Place p, prefix & name, 28 + (leftW - 10) / 2, h - 76, (leftW - 10) / 2, 34
    ButtonStyle p.Controls(prefix & name), False, (name = "btnDelete" Or name = "btnCancel")
    Place p, prefix & "btnNew", 18, h - 34, leftW, 28
End Sub

Private Sub HomePage(ByVal p As Object)
    Dim w As Single, h As Single, cw As Single, x As Single, i As Long, c As Object, titles As Variant, values As Variant
    Dim sideW As Single, mainX As Single, tileW As Single, tileH As Single
    w = mPageWidth: h = Application.Max(440, mPageHeight)
    titles = Array("Commandes historiques", "Produits", "Fournisseurs", "Produits sous seuil", "Appro. en attente")
    values = Array(Tbl("tblT_Commandes").ListRows.Count, Tbl("tblT_Produits").ListRows.Count, Tbl("tblT_Fournisseurs").ListRows.Count, ThisWorkbook.Worksheets("Menu").Range("E35").Value2, ThisWorkbook.Worksheets("Menu").Range("E36").Value2)
    cw = (w - 48 - 4 * 14) / 5
    For i = 0 To 4
        x = 18 + i * (cw + 14)
        Backdrop p, "homeCard" & i, x, 14, cw, 104, RGB(255, 255, 255)
        Set c = Heading(p, "homeMetricTitle" & i)
        c.Caption = titles(i): Bounds c, x + 6, 27, cw - 12, 28: TextStyle c, 11, True: c.TextAlign = 2
        Set c = Heading(p, "homeMetricValue" & i)
        c.Caption = Format$(values(i), "#,##0"): Bounds c, x + 6, 65, cw - 12, 37: TextStyle c, 27, True: c.TextAlign = 2
    Next i
    sideW = 214: mainX = sideW + 36
    Backdrop p, "homeSidebar", 18, 136, sideW, h - 154, RGB(255, 255, 255)
    Set c = Heading(p, "homeManagement")
    c.Caption = "Gestion": Bounds c, 30, 148, sideW - 24, 22: TextStyle c, 13, True
    For i = 1 To 5
        Place p, "nav" & i, 30, 182 + (i - 1) * 43, sideW - 24, 35
        ButtonStyle p.Controls("nav" & i)
    Next i
    p.Controls("nav3").Caption = "Commandes clients"
    p.Controls("nav4").Caption = "Approvisionnements"
    p.Controls("nav5").Caption = "Stocks et seuils"
    Place p, "lblWelcome", mainX, 136, w - mainX - 18, 31
    p.Controls("lblWelcome").Caption = "Gestion et analyse logistique": TextStyle p.Controls("lblWelcome"), 19, True
    Place p, "lblIntro", mainX, 176, w - mainX - 18, 32
    p.Controls("lblIntro").Caption = "Choisissez une rubrique pour ouvrir les outils et les rapports."
    TextStyle p.Controls("lblIntro"), 11, True: p.Controls("lblIntro").ForeColor = Muted()
    tileW = (w - mainX - 32) / 2: tileH = Application.Min(79, (h - 260) / 2)
    For i = 6 To 9
        x = mainX + ((i - 6) Mod 2) * (tileW + 14)
        Place p, "nav" & i, x, 224 + ((i - 6) \ 2) * (tileH + 16), tileW, tileH
        ButtonStyle p.Controls("nav" & i): p.Controls("nav" & i).Font.Size = 14
    Next i
    p.Controls("nav9").Caption = "Exports et aide"
    p.Controls("lblCounters").Visible = False
    Place p, "lblHomeState", mainX, h - 28, w - mainX - 18, 20
    p.Controls("lblHomeState").ForeColor = Muted()
End Sub

Private Sub SettingsPage(ByVal p As Object)
    Dim w As Single, h As Single, half As Single, i As Long, x As Single, y As Single
    w = mPageWidth: h = mPageHeight: half = (w - 48) / 2
    PageTitle p, "lblParamsTitle", w
    For i = 0 To 15
        x = 18 + (i \ 8) * (half + 12): y = 43 + (i Mod 8) * Application.Min(34, (h - 155) / 8)
        Place p, "lblParam" & i, x, y + 5, half * 0.43, 20
        Place p, "par" & i, x + half * 0.43, y, half * 0.57 - 4, 27
    Next i
    x = 18 + (half + 12): y = 43 + 7 * Application.Min(34, (h - 155) / 8)
    p.Controls("par15").Visible = False
    p.Controls("lblParam15").Caption = "Produit historique"
    TextStyle p.Controls("lblParam15"), 10.5, True
    Place p, "cboSimulationProduct", x + half * 0.43, y, half * 0.57 - 4, 27
    With p.Controls("cboSimulationProduct")
        .Visible = True
        .BackColor = RGB(255, 255, 255)
        .ForeColor = Ink()
        .Font.name = "Segoe UI"
        .Font.size = 11
        .SpecialEffect = 0
        .BorderStyle = 1
        .BorderColor = RGB(126, 154, 191)
    End With
    Place p, "btnSettingsSave", 18, h - 76, (w - 56) / 3, 34
    Place p, "btnSettingsReload", 28 + (w - 56) / 3, h - 76, (w - 56) / 3, 34
    Place p, "btnSettingsReset", 38 + 2 * (w - 56) / 3, h - 76, (w - 56) / 3, 34
    ButtonStyle p.Controls("btnSettingsSave"), True
    Place p, "lblSettingsNote", 18, h - 32, w - 36, 27
    p.Controls("lblSettingsNote").ForeColor = Muted()
End Sub

Private Sub SimulationPage(ByVal p As Object)
    Dim w As Single, h As Single, workW As Single, sideX As Single, i As Long, bw As Single, c As Object, stepY As Single, baseY As Single
    w = mPageWidth: h = mPageHeight: workW = w * 0.67 - 30: sideX = workW + 48
    PageTitle p, "lblSimulationTitle", w
    p.Controls("lblScenarioHeader").Visible = False
    Set c = Heading(p, "simulationHead0"): c.Caption = "Scenario": Bounds c, 18, 45, workW * 0.45, 22: TextStyle c, 10.5, True
    Set c = Heading(p, "simulationHead1"): c.Caption = "Seuil s": Bounds c, 18 + workW * 0.47, 45, workW * 0.23, 22: TextStyle c, 10.5, True
    Set c = Heading(p, "simulationHead2"): c.Caption = "Quantite Q": Bounds c, 18 + workW * 0.74, 45, workW * 0.26, 22: TextStyle c, 10.5, True
    For i = 1 To 3
        Place p, "scenarioName" & i, 18, 82 + (i - 1) * 39, workW * 0.45, 27
        Place p, "scenarios" & i, 18 + workW * 0.47, 77 + (i - 1) * 39, workW * 0.23, 28
        Place p, "scenarioq" & i, 18 + workW * 0.74, 77 + (i - 1) * 39, workW * 0.26, 28
    Next i
    bw = (workW - 12) / 2
    baseY = 196: stepY = Application.Min(48, (h - baseY - 82) / 2)
    Place p, "btnScenarioSave", 18, baseY, bw, 34
    Place p, "btnCustomResults", 30 + bw, baseY, bw, 34
    Place p, "btnCalibrate", 18, baseY + stepY, bw, 34
    Place p, "btnNormal", 30 + bw, baseY + stepY, bw, 34
    Place p, "btnCompare", 18, baseY + 2 * stepY, bw, 36
    Place p, "btnCustom", 30 + bw, baseY + 2 * stepY, bw, 36
    p.Controls("btnCompare").Caption = "Comparer et afficher les 3 scenarios"
    p.Controls("btnCustom").Caption = "Simuler les valeurs saisies s / Q"
    p.Controls("btnCustomResults").Caption = "Afficher les resultats personnalises"
    p.Controls("btnCalibrate").Caption = "Recalibrer et suggerer s / Q / S0"
    ButtonStyle p.Controls("btnCompare"), True
    ButtonStyle p.Controls("btnCustom"), True
    Place p, "lblDraws", sideX, 48, w - sideX - 18, 25
    TextStyle p.Controls("lblDraws"), 11, True
    Place p, "lstDraws", sideX, 79, w - sideX - 18, h - 134
    p.Controls("lstDraws").ColumnWidths = "65;120"
    Place p, "lblSimulationNote", 18, h - 43, workW, 35
    p.Controls("lblSimulationNote").Caption = "Suggestions s / Q et S0 adaptees au produit choisi dans Parametres. Valeurs modifiables. Recalibrer regenere les suggestions. Comparer ouvre Resultats."
    p.Controls("lblSimulationNote").ForeColor = Muted()
End Sub

Private Sub ResultsPage(ByVal p As Object)
    Dim w As Single, h As Single, bw As Single, unit As Single, i As Long, x As Single, cw As Single, headings As Variant, c As Object
    w = mPageWidth: h = mPageHeight
    PageTitle p, "lblResultsTitle", w
    Place p, "lblResultState", 18, 47, w - 36, 26
    TextStyle p.Controls("lblResultState"), 11, True
    Place p, "lblResultDecision", 18, 82, w - 36, 45
    p.Controls("lblResultsColumns").Visible = False
    Place p, "lstResults", 18, 182, w - 36, h - 282
    unit = (w - 212) / 7
    headings = Array("Scenario", "Cout ($)", "Service", "IC bas ($)", "IC haut ($)", "Stock moyen", "Ruptures", "Commandes")
    x = 18
    For i = 0 To 7
        If i = 0 Then cw = 160 Else cw = Int(unit)
        Set c = Heading(p, "resultsHead" & i)
        c.Caption = headings(i): Bounds c, x + 3, 147, cw - 3, 30
        TextStyle c, 10, True
        x = x + cw
    Next i
    p.Controls("lstResults").ColumnWidths = "160;" & Int(unit) & ";" & Int(unit) & ";" & Int(unit) & ";" & Int(unit) & ";" & Int(unit) & ";" & Int(unit) & ";" & Int(unit)
    bw = (w - 56) / 3
    Place p, "btnSummary", 18, h - 76, bw, 34
    Place p, "btnEvolution", 28 + bw, h - 76, bw, 34
    Place p, "btnResultsRefresh", 38 + bw * 2, h - 76, bw, 34
    p.Controls("btnSummary").Caption = "Ouvrir le rapport de synthese"
    p.Controls("btnEvolution").Caption = "Voir l'evolution du stock"
    p.Controls("btnResultsRefresh").Caption = "Actualiser les resultats"
    ButtonStyle p.Controls("btnSummary"), True
    ButtonStyle p.Controls("btnEvolution")
    p.Controls("btnEvolution").BackColor = RGB(247, 248, 250)
    p.Controls("btnEvolution").ForeColor = Ink()
    ButtonStyle p.Controls("btnResultsRefresh")
    p.Controls("btnResultsRefresh").BackColor = RGB(247, 248, 250)
    p.Controls("btnResultsRefresh").ForeColor = Ink()
    Place p, "lblResultsNote", 18, h - 31, w - 36, 24
    p.Controls("lblResultsNote").ForeColor = Muted()
End Sub

Private Sub ReportPage(ByVal p As Object)
    Dim w As Single, h As Single, i As Long, c As Object
    w = mPageWidth: h = mPageHeight
    Set c = Heading(p, "lblReportTitle")
    c.Caption = "Rapports de simulation": Bounds c, 18, 8, w - 36, 28
    TextStyle c, 17, True: c.ForeColor = Blue()
    Place p, "cboReportView", 18, 43, w - 36, 29
    Set c = Heading(p, "lblReportSubtitle")
    Bounds c, 18, 80, w - 36, 23: TextStyle c, 10.5
    Set c = Heading(p, "lblReportContext")
    Bounds c, 18, 108, w - 36, 35: TextStyle c, 10, True: c.ForeColor = Muted()
    Set c = p.Controls("lstReport")
    Bounds c, 18, 181, w - 36, h - 230
    c.IntegralHeight = False
    c.Font.name = "Segoe UI": c.Font.size = 10.5
    c.BackColor = RGB(255, 255, 255): c.ForeColor = Ink()
    c.BorderColor = RGB(205, 214, 228)
    For i = 0 To 7
        Set c = Heading(p, "reportHead" & i)
        TextStyle c, 10, True: c.ForeColor = BrandNavy()
    Next i
    Set c = Heading(p, "lblReportNotes")
    Bounds c, 18, h - 41, w - 36, 33: TextStyle c, 10
    c.ForeColor = Muted()
End Sub

Public Sub FitReportColumns(ByVal p As Object, ByVal widths As String, ByVal headers As Variant)
    Dim parts As Variant, total As Double, available As Single, x As Single, i As Long, cw As Long, used As Long, listWidths As String, c As Object
    parts = Split(widths, ";")
    For i = 0 To UBound(parts): total = total + Val(parts(i)): Next i
    available = p.Controls("lstReport").Width - 22: x = 18
    For i = 0 To 7
        Set c = p.Controls("reportHead" & i)
        c.Visible = (i <= UBound(parts))
        If c.Visible Then
            cw = Int(available * Val(parts(i)) / total)
            c.Caption = headers(i): Bounds c, x + 4, 148, cw - 4, 29
            If i > 0 Then listWidths = listWidths & ";"
            listWidths = listWidths & CStr(cw): x = x + cw: used = used + cw
        End If
    Next i
    p.Controls("lstReport").ColumnCount = UBound(parts) + 1
    p.Controls("lstReport").ColumnWidths = listWidths
End Sub

Public Sub Bounds(ByVal c As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
End Sub

Private Sub ExportPage(ByVal p As Object)
    Dim w As Single, h As Single
    w = mPageWidth: h = mPageHeight
    PageTitle p, "lblExportTitle", w
    Place p, "btnPowerBI", 18, 52, (w - 48) / 2, 38
    Place p, "btnExport", 30 + (w - 48) / 2, 52, (w - 48) / 2, 38
    ButtonStyle p.Controls("btnPowerBI"), True
    Place p, "lblGuideTitle", 18, 110, w - 36, 24
    TextStyle p.Controls("lblGuideTitle"), 12, True
    Place p, "txtGuide", 18, 141, w - 36, h - 196
    p.Controls("txtGuide").BackColor = RGB(255, 255, 255)
    Place p, "btnGuideSheet", 18, h - 43, 240, 34
End Sub

Public Sub LayoutDataCoForm(ByVal view As Object)
    Dim tabs As Object, p As Object, i As Long, active As Long, w As Single, h As Single, c As Object, nw As Single
    Dim stage As String
    On Error GoTo Failed
    stage = "Cadre"
    view.BackColor = Canvas()
    view.ScrollBars = 0: view.ScrollTop = 0: view.ScrollLeft = 0
    w = view.InsideWidth: h = view.InsideHeight
    Set tabs = view.Controls("tabsMain")
    active = tabs.Value
    Backdrop view, "lblBrandBar", 0, 0, w, 80, RGB(255, 255, 255)
    Backdrop view, "lblAccentBar", 0, 79, w, 2, Accent()
    Backdrop view, "lblFooterBar", 0, h - 47, w, 47, RGB(255, 255, 255)
    Place view, "lblAppTitle", 22, 12, w - 44, 34
    TextStyle view.Controls("lblAppTitle"), 25, True
    view.Controls("lblAppTitle").Caption = "DataCo Global"
    Place view, "lblAppSubtitle", 24, 51, w - 48, 20
    TextStyle view.Controls("lblAppSubtitle"), 11, True
    view.Controls("lblAppSubtitle").ForeColor = Muted()
    nw = (w - 44 - 10 * 5) / 11
    For i = 0 To 10
        Set c = view.Controls("designNav" & i)
        Bounds c, 22 + i * (nw + 5), 90, nw, 34
        ButtonStyle c: c.Font.Size = 10.5: c.Font.Italic = True
    Next i
    tabs.Style = 2: tabs.BackColor = Canvas()
    Bounds tabs, 12, 133, w - 24, h - 185
    mPageWidth = Application.Max(920, tabs.Width - 6)
    mPageHeight = Application.Max(408, tabs.Height - 6)
    For i = 0 To tabs.Pages.Count - 1
        stage = "Page " & i
        tabs.Value = i
        Set p = tabs.Pages(i)
        p.Zoom = 100
        p.ScrollBars = 0: p.KeepScrollBarsVisible = 0
        p.ScrollLeft = 0: p.ScrollTop = 0
        StylePage p
        Backdrop p, "designPageBackground", 0, 0, mPageWidth, mPageHeight, Canvas()
        Select Case i
            Case 0: HomePage p
            Case 1: CatalogPage p, "pr_", Array("txtKey", "txtName", "txtPrice"), Array("Reference", "Designation", "Prix ($)")
            Case 2: CatalogPage p, "fu_", Array("txtKey", "txtName", "txtDelay", "txtCost"), Array("Code", "Fournisseur", "Delai (j)", "Cout ($)")
            Case 3: CatalogPage p, "co_", Array("txtKey", "cboProduct", "txtQty", "txtDate", "txtClient"), Array("Commande", "Date", "Produit", "Quantite", "Statut")
            Case 4: CatalogPage p, "ap_", Array("txtKey", "cboSupplier", "cboProduct", "txtQty", "txtDate", "txtReceipt"), Array("Commande", "Fournisseur", "Produit", "Quantite", "Statut")
            Case 5: CatalogPage p, "st_", Array("cboProduct", "txtDelta", "txtThreshold"), Array("Produit", "Stock", "Seuil", "Etat")
            Case 6: SettingsPage p
            Case 7: SimulationPage p
            Case 8: ResultsPage p
            Case 9: ExportPage p
            Case 10: ReportPage p
        End Select
        p.Controls("designPageBackground").ZOrder 1
        p.ScrollWidth = mPageWidth: p.ScrollHeight = mPageHeight
        If i = 0 Then p.ScrollHeight = Application.Max(440, mPageHeight)
    Next i
    tabs.Value = active
    stage = "Pied de page"
    Place view, "designZoomOut", 22, h - 37, 38, 30
    Place view, "designZoomReset", 70, h - 37, 61, 30
    Place view, "designZoomIn", 141, h - 37, 38, 30
    Place view, "designZoomFit", 189, h - 37, 98, 30
    For Each c In view.Controls
        If Left$(c.Name, 10) = "designZoom" Then ButtonStyle c
    Next c
    Place view, "lblStatus", 306, h - 31, Application.Max(80, w - 574), 23
    TextStyle view.Controls("lblStatus"), 10.5, True
    view.Controls("lblStatus").ForeColor = Muted()
    Place view, "btnHome", w - 252, h - 37, 112, 30
    Place view, "btnClose", w - 128, h - 37, 112, 30
    ButtonStyle view.Controls("btnHome"): ButtonStyle view.Controls("btnClose")
    Exit Sub
Failed:
    Err.Raise Err.Number, "LayoutDataCoForm", stage & " : " & Err.Description
End Sub

Public Sub PaintDataCoNavigation(ByVal view As Object)
    Dim i As Long, c As Object
    For i = 0 To 10
        Set c = view.Controls("designNav" & i)
        c.BackColor = RGB(255, 255, 255): c.ForeColor = Ink()
        If i = view.Controls("tabsMain").Value Then
            c.BackColor = RGB(228, 232, 243): c.ForeColor = Blue()
        End If
    Next i
End Sub

Public Sub ZoomDataCoPages(ByVal view As Object, ByVal percent As Long, Optional ByVal pageIndex As Long = -1)
    Dim p As Object
    If pageIndex < 0 Then pageIndex = view.Controls("tabsMain").Value
    Set p = view.Controls("tabsMain").Pages(pageIndex)
    p.Zoom = percent
    p.ScrollBars = 3
    p.KeepScrollBarsVisible = 0
    p.ScrollTop = 0: p.ScrollLeft = 0
    view.Controls("designZoomReset").Caption = CStr(percent) & "%"
End Sub

Public Function FitDataCoZoom(ByVal view As Object) As Long
    Dim tabs As Object
    Set tabs = view.Controls("tabsMain")
    FitDataCoZoom = Application.Max(50, Application.Min(100, Int(Application.Min((tabs.Width - 12) / mPageWidth, (tabs.Height - 12) / tabs.SelectedItem.ScrollHeight) * 100)))
End Function







