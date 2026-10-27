Option Explicit

Public Type SimResult
    Cost As Double
    Service As Double
    lost As Long
    Orders As Long
    AvgStock As Double
    TotalDemand As Long
    TotalSales As Long
    Availability As Double
    ShortageDays As Long
    MaxShortage As Long
    HoldingCost As Double
    OrderingCost As Double
    ShortageCost As Double
    Coverage As Variant
End Type

Private rngState As Double
Public Function TestNewKPI() As String
    On Error GoTo Failed
    Dim d() As Long, journal() As Variant, result As SimResult, savedMu As Variant
    savedMu = Param("mu"): SetParam "mu", 10
    ReDim d(1 To 1, 1 To 6): ReDim journal(1 To 6, 1 To 7)
    d(1, 1) = 10: d(1, 2) = 10: d(1, 3) = 0: d(1, 4) = 10: d(1, 5) = 10: d(1, 6) = 10
    result = Simulate(d, 1, 6, 0, 0, 1, 20, False, 2, 3, 4, journal, True)
    Require result.ShortageDays = 5 And result.MaxShortage = 3, "Jours consecutifs de rupture incorrects."
    Require Abs(result.Availability - 1# / 6#) < 0.00000001, "Disponibilite incorrecte."
    Require result.lost = 50 And result.HoldingCost = 0 And result.OrderingCost = 3 And result.ShortageCost = 200 And result.Cost = 203, "Decomposition des couts incorrecte."
    Require result.Coverage = 0, "Couverture incorrecte."
    result = Simulate(d, 1, 6, 100, 0, 1, 1, True, 2, 3, 4, journal, True)
    Require result.ShortageDays = 0 And result.MaxShortage = 0 And result.Availability = 1, "Absence de rupture incorrecte."
    Require Abs(result.Coverage - result.AvgStock / 10#) < 0.00000001, "Couverture moyenne incorrecte."
    SetParam "mu", 0
    GenerateDemand d, 1, 6, 0, 0, 1
    result = Simulate(d, 1, 6, 0, 0, 1, 1, True, 0, 0, 0, journal, False)
    Require IsEmpty(result.Coverage) And result.Availability = 1 And result.Service = 1, "Demande nulle incorrecte."
    SetParam "mu", savedMu
    UpdateOperationalKPI
    Require ThisWorkbook.Worksheets("Menu").Range("E35").Value2 = 0 And ThisWorkbook.Worksheets("Menu").Range("E36").Value2 = 0, "KPI operationnels initiaux incorrects."
    AdjustStock "P001", -241, 260
    Require ThisWorkbook.Worksheets("Menu").Range("E35").Value2 = 1, "Produit sous seuil non compte : compteur=" & ThisWorkbook.Worksheets("Menu").Range("E35").Value2 & ", stock=" & Tbl("tblT_Stock").DataBodyRange.Cells(FindRow("tblT_Stock", "P001"), 2).Value2 & ", seuil=" & Tbl("tblT_Stock").DataBodyRange.Cells(FindRow("tblT_Stock", "P001"), 3).Value2
    AdjustStock "P001", 1, 260
    Require ThisWorkbook.Worksheets("Menu").Range("E35").Value2 = 0, "Stock egal au seuil compte comme alerte."
    Dim supply As String: supply = CreateSupply("F001", "P001", 10, Date)
    Require ThisWorkbook.Worksheets("Menu").Range("E36").Value2 = 1, "Approvisionnement en transit non compte."
    ReceiveSupply supply, Date + 4
    Require ThisWorkbook.Worksheets("Menu").Range("E36").Value2 = 0, "Approvisionnement recu toujours compte."
    Dim clientOrder As String
    clientOrder = CreateOrder("P001", 11, Date, "Test KPI")
    Require ThisWorkbook.Worksheets("Menu").Range("E35").Value2 = 1, "Alerte apres vente non actualisee."
    CancelOrder clientOrder
    Require ThisWorkbook.Worksheets("Menu").Range("E35").Value2 = 0, "Alerte apres annulation non actualisee."
    TestNewKPI = "PASS : disponibilite, jours, sequences, couts, couverture, mu nul, alertes et approvisionnements."
    Exit Function
Failed:
    SetParam "mu", savedMu
    TestNewKPI = "FAIL " & Err.Number & " : " & Err.Description
End Function

Public Function DiagnoseComparison() As String
    On Error GoTo Failed
    CompareScenarios True
    DiagnoseComparison = "PASS": Exit Function
Failed:
    DiagnoseComparison = CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub SeedRandom(ByVal seed As Long)
    rngState = Abs(CDbl(seed)) Mod 2147483646#
    If rngState < 1 Then rngState = 1
End Sub

Private Function UniformRandom() As Double
    Dim x As Double
    x = rngState * 16807#
    rngState = x - Fix(x / 2147483647#) * 2147483647#
    UniformRandom = rngState / 2147483647#
End Function

Public Function NormalDraw(ByVal mean As Double, ByVal sigma As Double) As Long
    Dim z As Double, x As Double
    z = Sqr(-2# * Log(UniformRandom())) * Cos(6.28318530717959 * UniformRandom())
    x = mean + sigma * z
    If x > 0 Then NormalDraw = CLng(Fix(x + 0.5)) Else NormalDraw = 0
End Function

Public Sub GenerateDemand(ByRef d() As Long, ByVal runs As Long, ByVal days As Long, ByVal mean As Double, ByVal sigma As Double, ByVal seed As Long)
    Dim r As Long, t As Long
    ReDim d(1 To runs, 1 To days): SeedRandom seed
    For r = 1 To runs: For t = 1 To days: d(r, t) = NormalDraw(mean, sigma): Next t: Next r
End Sub

Public Function Simulate(ByRef d() As Long, ByVal run As Long, ByVal days As Long, ByVal initial As Long, ByVal threshold As Long, ByVal quantity As Long, ByVal delay As Long, ByVal multiple As Boolean, ByVal holding As Double, ByVal orderCost As Double, ByVal lostCost As Double, ByRef journal() As Variant, ByVal writeJournal As Boolean, Optional ByVal code As Long = 1, Optional ByVal rowOffset As Long = 0) As SimResult
    Dim arrivals() As Long, t As Long, stock As Long, transit As Long, sales As Long, lost As Long, ordered As Long, sumStock As Double, result As SimResult, streak As Long
    ReDim arrivals(1 To days + delay)
    stock = initial
    For t = 1 To days
        stock = stock + arrivals(t): transit = transit - arrivals(t)
        sales = d(run, t): If sales > stock Then sales = stock
        lost = d(run, t) - sales: stock = stock - sales
        If lost > 0 Then
            result.ShortageDays = result.ShortageDays + 1: streak = streak + 1
            If streak > result.MaxShortage Then result.MaxShortage = streak
        Else
            streak = 0
        End If
        ordered = 0
        If stock + transit <= threshold And (multiple Or transit = 0) Then
            ordered = quantity: arrivals(t + delay) = arrivals(t + delay) + quantity
            transit = transit + quantity: result.Orders = result.Orders + 1
        End If
        result.TotalDemand = result.TotalDemand + d(run, t)
        result.TotalSales = result.TotalSales + sales: result.lost = result.lost + lost
        sumStock = sumStock + stock
        If writeJournal Then
            journal(rowOffset + t, 1) = code: journal(rowOffset + t, 2) = t
            journal(rowOffset + t, 3) = d(run, t): journal(rowOffset + t, 4) = stock
            journal(rowOffset + t, 5) = sales: journal(rowOffset + t, 6) = lost: journal(rowOffset + t, 7) = ordered
        End If
    Next t
    result.AvgStock = sumStock / days
    If result.TotalDemand > 0 Then result.Service = result.TotalSales / result.TotalDemand Else result.Service = 1
    result.Availability = (days - result.ShortageDays) / days
    result.HoldingCost = sumStock * holding
    result.OrderingCost = result.Orders * orderCost
    result.ShortageCost = result.lost * lostCost
    result.Cost = result.HoldingCost + result.OrderingCost + result.ShortageCost
    result.Coverage = Empty
    If CDbl(Param("mu")) > 0 Then result.Coverage = result.AvgStock / CDbl(Param("mu"))
    Simulate = result
End Function

Public Sub ValidateParameters()
    Dim v As Double
    v = NumberValue(Param("S0"), "S0", 0, True): Require v <= 1000000, "S0 trop grand."
    v = NumberValue(Param("s"), "s", 0, True): Require v <= 1000000, "s trop grand."
    v = NumberValue(Param("Q"), "Q", 1, True): Require v <= 1000000, "Q trop grand."
    v = NumberValue(Param("L"), "L", 1, True): Require v <= 3650, "L doit etre <= 3650."
    v = NumberValue(Param("mu"), "mu", 0): Require v <= 1000000, "mu trop grand."
    v = NumberValue(Param("sigma"), "sigma", 0): Require v <= 1000000, "sigma trop grand."
    v = NumberValue(Param("Cpos"), "Cpos", 0)
    v = NumberValue(Param("Ccmd"), "Ccmd", 0)
    v = NumberValue(Param("Crup"), "Crup", 0)
    v = NumberValue(Param("N"), "N", 1, True): Require v <= 3650, "N doit etre <= 3650."
    v = NumberValue(Param("R"), "R", 2, True): Require v <= 1000, "R doit etre <= 1000."
    v = NumberValue(Param("Graine"), "Graine", 1, True): Require v <= 2147483646#, "Graine hors limites."
    v = NumberValue(Param("SeuilService"), "SeuilService", 0): Require v <= 1, "SeuilService entre 0 et 1."
    Require Param("CommandesMultiples") = "Oui" Or Param("CommandesMultiples") = "Non", "CommandesMultiples : Oui ou Non."
    Require FindRow("tblT_Produits", CStr(Param("ProduitSimulation"))) > 0, "Le produit de simulation doit exister dans le catalogue."
    Require CDbl(Param("N")) * CDbl(Param("mu")) < 1000000000#, "Demande cumulee trop grande."
End Sub

Private Sub MeanSD(ByRef x() As Double, ByVal n As Long, ByRef mean As Double, ByRef sd As Double)
    Dim i As Long, ss As Double
    mean = 0: sd = 0
    For i = 1 To n: mean = mean + x(i): Next i
    mean = mean / n
    For i = 1 To n: ss = ss + (x(i) - mean) ^ 2: Next i
    sd = Sqr(ss / (n - 1))
End Sub

Public Sub CompareScenarios(Optional ByVal silent As Boolean = False)
    On Error GoTo Failed
    Dim days As Long, runs As Long, d() As Long, journal() As Variant, kpi() As Variant, sens() As Variant
    Dim stats() As Variant, pairs() As Variant, decision() As Variant
    Dim costs() As Double, services() As Double, x() As Double, lostSum(1 To 3) As Double
    Dim stockSum(1 To 3) As Double, ordersSum(1 To 3) As Double, serviceMeans(1 To 3) As Double
    Dim mean As Double, sd As Double, crit As Double, best As Long, bestCost As Double, count As Long
    Dim sc As Variant, result As SimResult, c As Long, r As Long, mode As Long, i As Long, j As Long, k As Long
    Dim multiple As Boolean, prevCalc As XlCalculation, prevEvents As Boolean, prevScreen As Boolean, stateSaved As Boolean, decisionText As String, phase As String
    EnsureWorkbookVisible
    phase = "validation"
    ValidateParameters
    days = CLng(Param("N")): runs = CLng(Param("R")): multiple = (Param("CommandesMultiples") = "Oui")
    phase = "scenarios"
    sc = Tbl("tblScenarios").DataBodyRange.Value2
    Require UBound(sc, 1) = 3, "Trois scenarios sont necessaires."
    For c = 1 To 3
        Require sc(c, 1) = c, "Codes des scenarios attendus : 1, 2, 3."
        Require NumberValue(sc(c, 4), "Seuil scenario", 0, True) <= 1000000, "Seuil trop grand."
        Require NumberValue(sc(c, 5), "Q scenario", 1, True) <= 1000000, "Q trop grand."
    Next c
    prevCalc = Application.Calculation: prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating: stateSaved = True
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    phase = "demandes"
    GenerateDemand d, runs, days, CDbl(Param("mu")), CDbl(Param("sigma")), CLng(Param("Graine"))
    ReDim journal(1 To 3 * days, 1 To 7): ReDim kpi(1 To 3 * runs, 1 To 15)
    ReDim sens(1 To 6 * runs, 1 To 9): ReDim costs(1 To 3, 1 To runs): ReDim services(1 To 3, 1 To runs)
    phase = "simulation"
    For c = 1 To 3
        For r = 1 To runs
            result = Simulate(d, r, days, CLng(Param("S0")), CLng(sc(c, 4)), CLng(sc(c, 5)), CLng(Param("L")), multiple, CDbl(Param("Cpos")), CDbl(Param("Ccmd")), CDbl(Param("Crup")), journal, r = 1, c, (c - 1) * days)
            i = (c - 1) * runs + r
            kpi(i, 1) = c: kpi(i, 2) = r: kpi(i, 3) = result.Cost: kpi(i, 4) = result.Service
            kpi(i, 5) = 1 - result.Service: kpi(i, 6) = result.Orders: kpi(i, 7) = result.AvgStock: kpi(i, 8) = result.lost
            kpi(i, 9) = result.Availability: kpi(i, 10) = result.ShortageDays: kpi(i, 11) = result.MaxShortage
            kpi(i, 12) = result.HoldingCost: kpi(i, 13) = result.OrderingCost: kpi(i, 14) = result.ShortageCost: kpi(i, 15) = result.Coverage
            costs(c, r) = result.Cost: services(c, r) = result.Service
            lostSum(c) = lostSum(c) + result.lost: stockSum(c) = stockSum(c) + result.AvgStock: ordersSum(c) = ordersSum(c) + result.Orders
            For mode = 1 To 2
                result = Simulate(d, r, days, CLng(Param("S0")), CLng(sc(c, 4)), CLng(sc(c, 5)), CLng(Param("L")), mode = 1, CDbl(Param("Cpos")), CDbl(Param("Ccmd")), CDbl(Param("Crup")), journal, False)
                j = (c - 1) * runs * 2 + (mode - 1) * runs + r
                sens(j, 1) = c: sens(j, 2) = r: sens(j, 3) = IIf(mode = 1, "Oui", "Non")
                sens(j, 4) = result.Cost: sens(j, 5) = result.Service: sens(j, 6) = 1 - result.Service
                sens(j, 7) = result.Orders: sens(j, 8) = result.AvgStock: sens(j, 9) = result.lost
            Next mode
        Next r
    Next c
    phase = "statistiques"
    crit = Application.WorksheetFunction.T_Inv_2T(0.05, runs - 1)
    ReDim stats(1 To 3, 1 To 12): ReDim x(1 To runs): bestCost = 1E+100
    For c = 1 To 3
        For r = 1 To runs: x(r) = costs(c, r): Next r
        MeanSD x, runs, mean, sd
        stats(c, 1) = c: stats(c, 2) = mean: stats(c, 3) = sd
        stats(c, 4) = mean - crit * sd / Sqr(runs): stats(c, 5) = mean + crit * sd / Sqr(runs)
        For r = 1 To runs: x(r) = services(c, r): Next r
        MeanSD x, runs, mean, sd: serviceMeans(c) = mean
        stats(c, 6) = mean: stats(c, 7) = sd: stats(c, 8) = mean - crit * sd / Sqr(runs): stats(c, 9) = mean + crit * sd / Sqr(runs)
        stats(c, 10) = lostSum(c) / runs: stats(c, 11) = ordersSum(c) / runs: stats(c, 12) = stockSum(c) / runs
        If mean >= CDbl(Param("SeuilService")) And stats(c, 2) < bestCost Then best = c: bestCost = stats(c, 2)
    Next c
    ReDim pairs(1 To 3, 1 To 7): k = 0
    For i = 1 To 2: For j = i + 1 To 3
        k = k + 1
        For r = 1 To runs: x(r) = costs(i, r) - costs(j, r): Next r
        MeanSD x, runs, mean, sd
        pairs(k, 1) = i: pairs(k, 2) = j: pairs(k, 3) = mean: pairs(k, 4) = sd
        pairs(k, 5) = mean - crit * sd / Sqr(runs): pairs(k, 6) = mean + crit * sd / Sqr(runs)
        If pairs(k, 5) <= 0 And pairs(k, 6) >= 0 Then
            pairs(k, 7) = "Equivalents statistiquement"
        ElseIf pairs(k, 6) < 0 Then
            pairs(k, 7) = "A moins couteux"
        Else
            pairs(k, 7) = "B moins couteux"
        End If
    Next j: Next i
    If best = 0 Then
        decisionText = "Aucun scenario n'atteint le seuil de service."
    Else
        decisionText = sc(best, 2)
        For k = 1 To 3
            If pairs(k, 5) <= 0 And pairs(k, 6) >= 0 Then
                i = pairs(k, 1): j = pairs(k, 2)
                If i = best And serviceMeans(j) >= CDbl(Param("SeuilService")) Then decisionText = decisionText & " / " & sc(j, 2) & " (equivalents)"
                If j = best And serviceMeans(i) >= CDbl(Param("SeuilService")) Then decisionText = decisionText & " / " & sc(i, 2) & " (equivalents)"
            End If
        Next k
    End If
    ReDim decision(1 To 1, 1 To 3): decision(1, 1) = decisionText: decision(1, 2) = CDbl(Param("SeuilService")): decision(1, 3) = CDbl(Now)
    phase = "exports"
    WriteTable "tblExport_PowerBI", journal, 3 * days, 7
    WriteTable "tblExport_KPI", kpi, 3 * runs, 15
    WriteTable "tblExport_Sensibilite", sens, 6 * runs, 9
    WriteTable "tblExport_Stats", stats, 3, 12: WriteTable "tblExport_Apparie", pairs, 3, 7
    WriteTable "tblExport_Decision", decision, 1, 3
    phase = "synthese"
    UpdateSummary days
    ThisWorkbook.Worksheets("Menu").Range("B30").Value2 = "Resultats actualises : " & Format$(Now, "dd/mm/yyyy hh:nn")
    ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2 = "ACTUEL"
    ThisWorkbook.Worksheets("Synthese").Calculate
    Application.Calculation = prevCalc: Application.EnableEvents = prevEvents: Application.ScreenUpdating = prevScreen
    If Not silent Then F_Main.ShowScenarioResults
    Exit Sub
Failed:
    Dim errorCode As Long, message As String: errorCode = Err.Number: message = phase & " : " & Err.Description
    If stateSaved Then
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
    End If
    MarkStale
    If silent Then Err.Raise errorCode, , message Else MsgBox message, vbExclamation
End Sub

Public Sub UpdateSummary(ByVal days As Long)
    Dim s As Worksheet, i As Long, t As Long, a() As Variant, journal As Variant
    Set s = ThisWorkbook.Worksheets("Synthese")
    s.Range("B5").Formula = "=Export_Decision!A2"
    For i = 1 To 3
        s.Cells(i + 7, 2).Formula = "=Scenarios!B" & i + 1
        s.Cells(i + 7, 3).Formula = "=Export_Stats!B" & i + 1
        s.Cells(i + 7, 4).Formula = "=Export_Stats!F" & i + 1
        s.Cells(i + 7, 5).Formula = "=Export_Stats!D" & i + 1
        s.Cells(i + 7, 6).Formula = "=Export_Stats!E" & i + 1
        s.Cells(i + 7, 7).Formula = "=Export_Stats!J" & i + 1
        s.Cells(i + 7, 8).Formula = "=Export_Stats!K" & i + 1
        s.Cells(i + 7, 9).Formula = "=Export_Stats!L" & i + 1
    Next i
    s.Range("C8:C10").NumberFormat = "$#,##0.00": s.Range("D8:D10").NumberFormat = "0.00%"
    s.Range("E8:F10").NumberFormat = "$#,##0.00": s.Range("G8:I10").NumberFormat = "#,##0.0"
    s.Range("B14:H14").value = Array("Scenario A", "Scenario B", "Difference ($)", "Ecart-type", "IC bas", "IC haut", "Conclusion")
    s.Range("B15:H17").Value2 = Tbl("tblExport_Apparie").DataBodyRange.Value2
    s.Range("D15:G17").NumberFormat = "0.00"
    s.Range("B21:J21").value = Array("Scenario", "Disponibilite", "Jours rupture", "Rupture max (j)", "Couverture (j)", "Objectif atteint", "Possession ($)", "Commande ($)", "Rupture ($)")
    For i = 1 To 3
        s.Cells(i + 21, 2).Formula = "=Scenarios!B" & i + 1
        s.Cells(i + 21, 3).Formula = "=AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[TauxDispo])"
        s.Cells(i + 21, 4).Formula = "=AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[JoursRupture])"
        s.Cells(i + 21, 5).Formula = "=AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[RuptureMax])"
        s.Cells(i + 21, 6).Formula = "=IF(Feuille_Controle!C10=0,"""",AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[Couverture]))"
        s.Cells(i + 21, 7).Formula = "=COUNTIFS(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[TauxService],"">=""&Feuille_Controle!C18)/COUNTIF(tblExport_KPI[CodeScenario]," & i & ")"
        s.Cells(i + 21, 8).Formula = "=AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[CoutPossession])"
        s.Cells(i + 21, 9).Formula = "=AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[CoutCommande])"
        s.Cells(i + 21, 10).Formula = "=AVERAGEIF(tblExport_KPI[CodeScenario]," & i & ",tblExport_KPI[CoutRupture])"
    Next i
    s.Range("C22:C24").NumberFormat = "0.00%": s.Range("G22:G24").NumberFormat = "0.00%"
    s.Range("D22:F24").NumberFormat = "0.00": s.Range("H22:J24").NumberFormat = "$#,##0.00"
    journal = Tbl("tblExport_PowerBI").DataBodyRange.Value2
    ReDim a(1 To days, 1 To 4)
    For t = 1 To days
        a(t, 1) = t
        For i = 1 To 3: a(t, i + 1) = journal((i - 1) * days + t, 4): Next i
    Next t
    With ThisWorkbook.Worksheets("Evolution_Stock")
        .Range("B6:E3656").ClearContents: .Range("B6").Resize(days, 4).Value2 = a
        If .ChartObjects.count > 0 Then
            Dim chart As chart, series As series
            Set chart = .ChartObjects(1).chart
            Do While chart.SeriesCollection.count > 0: chart.SeriesCollection(1).Delete: Loop
            For i = 1 To 3
                Set series = chart.SeriesCollection.NewSeries
                series.name = CStr(.Cells(5, i + 2).Value2)
                series.XValues = .Range("B6:B" & days + 5)
                series.values = .Cells(6, i + 2).Resize(days, 1)
                series.Format.line.Weight = 1.25
                If i = 1 Then series.Format.line.ForeColor.RGB = RGB(37, 99, 235)
                If i = 2 Then series.Format.line.ForeColor.RGB = RGB(217, 119, 6)
                If i = 3 Then series.Format.line.ForeColor.RGB = RGB(5, 150, 105)
            Next i
            chart.Legend.Position = xlLegendPositionBottom
            chart.Axes(xlCategory).TickLabelSpacing = 30
            chart.Axes(xlCategory).TickLabels.Orientation = 0
        End If
    End With
End Sub

Public Sub TestNormal()
    Dim i As Long, a(1 To 10, 1 To 2) As Variant
    SeedRandom CLng(Param("Graine"))
    For i = 1 To 10: a(i, 1) = i: a(i, 2) = NormalDraw(CDbl(Param("mu")), CDbl(Param("sigma"))): Next i
    ThisWorkbook.Worksheets("Feuille_Controle").Range("F5:G5").value = Array("Tirage", "Demande")
    ThisWorkbook.Worksheets("Feuille_Controle").Range("F6:G15").Value2 = a
End Sub

Public Sub RunCustom(Optional ByVal silent As Boolean = False)
    On Error GoTo Failed
    EnsureWorkbookVisible
    ValidateParameters
    Dim d() As Long, j() As Variant, result As SimResult, runs As Long, days As Long, r As Long, s As Worksheet
    runs = CLng(Param("R")): days = CLng(Param("N"))
    GenerateDemand d, runs, days, CDbl(Param("mu")), CDbl(Param("sigma")), CLng(Param("Graine"))
    ReDim j(1 To days, 1 To 7)
    Set s = ThisWorkbook.Worksheets("Feuille_Controle")
    s.Range("F18:J1018").ClearContents: s.Range("F18:J18").value = Array("Run", "Cout ($)", "Service", "Ruptures", "Commandes")
    Dim a() As Variant: ReDim a(1 To runs, 1 To 5)
    For r = 1 To runs
        result = Simulate(d, r, days, CLng(Param("S0")), CLng(Param("s")), CLng(Param("Q")), CLng(Param("L")), Param("CommandesMultiples") = "Oui", CDbl(Param("Cpos")), CDbl(Param("Ccmd")), CDbl(Param("Crup")), j, False)
        a(r, 1) = r: a(r, 2) = result.Cost: a(r, 3) = result.Service: a(r, 4) = result.lost: a(r, 5) = result.Orders
    Next r
    s.Range("F19").Resize(runs, 5).Value2 = a: s.Range("H19:H1018").NumberFormat = "0.00%"
    s.Range("G19:G1018").NumberFormat = "$#,##0.00"
    s.Range("L18").Value2 = "Produit / s / Q / N / R / L / S0 / mu / sigma / couts / graine / multiples"
    s.Range("L19").Value2 = CustomSignature()
    Exit Sub
Failed:
    If silent Then Err.Raise Err.Number, Err.source, Err.Description Else MsgBox Err.Description, vbExclamation
End Sub


Public Function CustomSignature() As String
    Dim key As Variant
    For Each key In Array("ProduitSimulation", "s", "Q", "N", "R", "L", "S0", "mu", "sigma", "Cpos", "Ccmd", "Crup", "Graine", "CommandesMultiples")
        CustomSignature = CustomSignature & CStr(Param(CStr(key))) & ";"
    Next key
End Function
