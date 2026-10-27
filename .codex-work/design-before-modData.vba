Option Explicit

Public Function Tbl(ByVal tableName As String) As ListObject
    Dim sh As Worksheet, lo As ListObject
    For Each sh In ThisWorkbook.Worksheets
        For Each lo In sh.ListObjects
            If lo.name = tableName Then Set Tbl = lo: Exit Function
        Next lo
    Next sh
    Err.Raise vbObjectError + 1, , "Table introuvable : " & tableName
End Function

Public Function Param(ByVal key As String) As Variant
    Dim i As Long
    For i = 6 To 21
        If ThisWorkbook.Worksheets("Feuille_Controle").Cells(i, 2).Value2 = key Then
            Param = ThisWorkbook.Worksheets("Feuille_Controle").Cells(i, 3).Value2: Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2, , "Parametre introuvable : " & key
End Function

Public Sub SetParam(ByVal key As String, ByVal value As Variant)
    Dim i As Long
    For i = 6 To 21
        If ThisWorkbook.Worksheets("Feuille_Controle").Cells(i, 2).Value2 = key Then
            ThisWorkbook.Worksheets("Feuille_Controle").Cells(i, 3).Value2 = value: Exit Sub
        End If
    Next i
End Sub

Public Function FindRow(ByVal name As String, ByVal key As String) As Long
    Dim a As Variant, i As Long, lo As ListObject
    Set lo = Tbl(name)
    If lo.DataBodyRange Is Nothing Then Exit Function
    a = lo.DataBodyRange.Value2
    For i = 1 To UBound(a, 1)
        If CStr(a(i, 1)) = key Then FindRow = i: Exit Function
    Next i
End Function

Public Sub AppendRow(ByVal tableName As String, ByVal values As Variant)
    Dim a() As Variant, i As Long, row As ListRow, lo As ListObject
    ReDim a(1 To 1, 1 To UBound(values) - LBound(values) + 1)
    For i = LBound(values) To UBound(values): a(1, i - LBound(values) + 1) = values(i): Next i
    Set lo = Tbl(tableName)
    lo.Parent.Unprotect
    Set row = lo.ListRows.Add
    row.Range.Value2 = a
    lo.Parent.Protect UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Public Sub WriteTable(ByVal name As String, ByRef a() As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    On Error GoTo Failed
    Dim lo As ListObject
    Set lo = Tbl(name)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Dim target As Range
    Set target = lo.Parent.Cells(lo.Range.row, lo.Range.Column)
    If rowCount > 0 Then target.Offset(1, 0).Resize(rowCount, colCount).Value2 = a
    lo.Resize target.Resize(rowCount + 1, colCount)
    Exit Sub
Failed:
    Err.Raise Err.Number, , name & " : " & Err.Description
End Sub

Public Function NextId(ByVal tableName As String, ByVal prefix As String) As String
    Dim a As Variant, i As Long, m As Long, k As String
    If Not Tbl(tableName).DataBodyRange Is Nothing Then
        a = Tbl(tableName).DataBodyRange.Value2
        For i = 1 To UBound(a, 1)
            k = CStr(a(i, 1))
            If Left$(k, Len(prefix)) = prefix Then
                If Val(Mid$(k, Len(prefix) + 1)) > m Then m = CLng(Val(Mid$(k, Len(prefix) + 1)))
            End If
        Next i
    End If
    NextId = prefix & Format$(m + 1, "000000")
End Function

Public Sub Require(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise vbObjectError + 10, , message
End Sub

Public Function NumberValue(ByVal value As Variant, ByVal label As String, Optional ByVal minimum As Double = 0, Optional ByVal integerOnly As Boolean = False) As Double
    Require Len(Trim$(CStr(value))) > 0, label & " obligatoire."
    Require IsNumeric(value), label & " doit etre numerique."
    Dim n As Double: n = CDbl(value)
    Require n >= minimum, label & " doit etre >= " & minimum & "."
    If integerOnly Then Require n = Fix(n), label & " doit etre entier."
    NumberValue = n
End Function

Public Function ParseDate(ByVal value As String) As Date
    Dim a As Variant, dt As Date
    a = Split(value, "/")
    Require UBound(a) = 2, "Date attendue : jj/mm/aaaa."
    Require Len(a(2)) = 4 And IsNumeric(a(0)) And IsNumeric(a(1)) And IsNumeric(a(2)), "Date invalide."
    dt = DateSerial(CLng(a(2)), CLng(a(1)), CLng(a(0)))
    Require Day(dt) = CLng(a(0)) And Month(dt) = CLng(a(1)) And Year(dt) = CLng(a(2)), "Date invalide."
    ParseDate = dt
End Function

Public Sub MarkStale()
    ThisWorkbook.Worksheets("Menu").Range("B30").Value2 = "Donnees modifiees : relancer la comparaison."
    ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2 = "PERIME"
End Sub

Public Function Referenced(ByVal tableName As String, ByVal col As Long, ByVal key As String) As Boolean
    Dim a As Variant, i As Long
    If Tbl(tableName).DataBodyRange Is Nothing Then Exit Function
    a = Tbl(tableName).DataBodyRange.Value2
    For i = 1 To UBound(a, 1)
        If CStr(a(i, col)) = key Then Referenced = True: Exit Function
    Next i
End Function

Public Sub SaveProduct(ByVal ref As String, ByVal designation As String, ByVal price As Double)
    Dim i As Long
    ref = UCase$(Trim$(ref)): designation = Trim$(designation)
    Require ref <> "" And designation <> "", "Reference et designation obligatoires."
    Require price >= 0, "Prix negatif interdit."
    i = FindRow("tblT_Produits", ref)
    If i = 0 Then
        AppendRow "tblT_Produits", Array(ref, designation, price)
        AppendRow "tblT_Stock", Array(ref, 0, 260, "ALERTE")
    Else
        Tbl("tblT_Produits").DataBodyRange.Cells(i, 2).Value2 = designation
        Tbl("tblT_Produits").DataBodyRange.Cells(i, 3).Value2 = price
    End If
    Call RefreshStocks: Call MarkStale
End Sub

Public Sub DeleteProduct(ByVal ref As String)
    Dim i As Long, j As Long
    i = FindRow("tblT_Produits", ref): Require i > 0, "Produit introuvable."
    Require ref <> "P001", "Le produit de simulation P001 doit etre conserve."
    Require Not Referenced("tblT_Commandes", 4, ref), "Produit utilise dans des commandes."
    Require Not Referenced("tblT_Approvisionnements", 4, ref), "Produit utilise dans des approvisionnements."
    Require Not Referenced("tblT_Mouvements", 3, ref), "Produit utilise dans des mouvements."
    j = FindRow("tblT_Stock", ref)
    If j > 0 Then
        Require Tbl("tblT_Stock").DataBodyRange.Cells(j, 2).Value2 = 0, "Le stock doit etre nul."
        DeleteTableRow "tblT_Stock", j
    End If
    DeleteTableRow "tblT_Produits", i: MarkStale
End Sub

Public Sub SaveSupplier(ByVal code As String, ByVal name As String, ByVal delay As Double, ByVal Cost As Double)
    Dim i As Long
    code = UCase$(Trim$(code)): name = Trim$(name)
    Require code <> "" And name <> "", "Code et nom obligatoires."
    Require delay >= 0 And Cost >= 0, "Delai et cout doivent etre positifs ou nuls."
    i = FindRow("tblT_Fournisseurs", code)
    If i = 0 Then
        AppendRow "tblT_Fournisseurs", Array(code, name, delay, Cost, Empty, Empty, "Personnalise")
    Else
        With Tbl("tblT_Fournisseurs").DataBodyRange
            .Cells(i, 2).Value2 = name: .Cells(i, 3).Value2 = delay: .Cells(i, 4).Value2 = Cost
        End With
    End If
    MarkStale
End Sub

Public Sub DeleteSupplier(ByVal code As String)
    Dim i As Long: i = FindRow("tblT_Fournisseurs", code)
    Require i > 0, "Fournisseur introuvable."
    Require Not Referenced("tblT_Commandes", 10, code), "Fournisseur utilise dans l'historique."
    Require Not Referenced("tblT_Approvisionnements", 3, code), "Fournisseur utilise dans des approvisionnements."
    DeleteTableRow "tblT_Fournisseurs", i: MarkStale
End Sub

Public Sub DeleteTableRow(ByVal name As String, ByVal i As Long)
    Dim lo As ListObject: Set lo = Tbl(name)
    lo.Parent.Unprotect
    lo.ListRows(i).Delete
    lo.Parent.Protect UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Public Function CreateOrder(ByVal product As String, ByVal qty As Long, ByVal dt As Date, ByVal client As String) As String
    Dim i As Long, stock As Long, sales As Long, key As String
    Require FindRow("tblT_Produits", product) > 0, "Produit introuvable."
    Require qty > 0 And Trim$(client) <> "", "Quantite positive et client obligatoires."
    i = FindRow("tblT_Stock", product): stock = Tbl("tblT_Stock").DataBodyRange.Cells(i, 2).Value2
    sales = qty: If sales > stock Then sales = stock
    key = NextId("tblT_Commandes", "C")
    AppendRow "tblT_Commandes", Array(key, CDbl(dt), Trim$(client), product, qty, sales, qty - sales, "Validee", "Saisie", "", Empty, Empty, "", Empty, Empty, "Demonstration", "", "")
    Tbl("tblT_Stock").DataBodyRange.Cells(i, 2).Value2 = stock - sales
    If sales > 0 Then LogMovement product, "Vente", -sales, key
    Call RefreshStocks: Call MarkStale: CreateOrder = key
End Function

Public Sub CancelOrder(ByVal key As String)
    Dim i As Long, j As Long, product As String, sales As Long
    i = FindRow("tblT_Commandes", key): Require i > 0, "Commande introuvable."
    With Tbl("tblT_Commandes").DataBodyRange
        Require .Cells(i, 9).Value2 = "Saisie", "L'historique DataCo est conserve."
        Require .Cells(i, 8).Value2 <> "Annulee", "Commande deja annulee."
        product = .Cells(i, 4).Value2: sales = .Cells(i, 6).Value2
        .Cells(i, 8).Value2 = "Annulee"
    End With
    j = FindRow("tblT_Stock", product)
    Tbl("tblT_Stock").DataBodyRange.Cells(j, 2).Value2 = Tbl("tblT_Stock").DataBodyRange.Cells(j, 2).Value2 + sales
    If sales > 0 Then LogMovement product, "Annulation", sales, key
    Call RefreshStocks: Call MarkStale
End Sub

Public Function CreateSupply(ByVal supplier As String, ByVal product As String, ByVal qty As Long, ByVal dt As Date) As String
    Dim key As String, i As Long, delay As Double
    i = FindRow("tblT_Fournisseurs", supplier): Require i > 0, "Fournisseur introuvable."
    Require FindRow("tblT_Produits", product) > 0 And qty > 0, "Produit et quantite invalides."
    delay = Tbl("tblT_Fournisseurs").DataBodyRange.Cells(i, 3).Value2
    key = NextId("tblT_Approvisionnements", "A")
    AppendRow "tblT_Approvisionnements", Array(key, CDbl(dt), supplier, product, qty, CDbl(dt) + delay, Empty, "En transit")
    Call UpdateOperationalKPI: Call MarkStale: CreateSupply = key
End Function

Public Sub ReceiveSupply(ByVal key As String, ByVal dt As Date)
    Dim i As Long, j As Long, product As String, qty As Long
    i = FindRow("tblT_Approvisionnements", key): Require i > 0, "Approvisionnement introuvable."
    With Tbl("tblT_Approvisionnements").DataBodyRange
        Require .Cells(i, 8).Value2 = "En transit", "Reception deja enregistree."
        Require CDbl(dt) >= .Cells(i, 2).Value2, "Reception anterieure a la commande."
        product = .Cells(i, 4).Value2: qty = .Cells(i, 5).Value2
        .Cells(i, 7).Value2 = CDbl(dt): .Cells(i, 8).Value2 = "Recue"
    End With
    j = FindRow("tblT_Stock", product)
    Tbl("tblT_Stock").DataBodyRange.Cells(j, 2).Value2 = Tbl("tblT_Stock").DataBodyRange.Cells(j, 2).Value2 + qty
    LogMovement product, "Reception", qty, key: RefreshStocks: MarkStale
End Sub

Public Sub AdjustStock(ByVal product As String, ByVal delta As Long, ByVal threshold As Long)
    Dim i As Long, stock As Long
    i = FindRow("tblT_Stock", product): Require i > 0, "Produit introuvable."
    stock = Tbl("tblT_Stock").DataBodyRange.Cells(i, 2).Value2
    Require stock + delta >= 0 And threshold >= 0, "Stock final ou seuil negatif."
    Tbl("tblT_Stock").DataBodyRange.Cells(i, 2).Value2 = stock + delta
    Tbl("tblT_Stock").DataBodyRange.Cells(i, 3).Value2 = threshold
    If delta <> 0 Then LogMovement product, "Ajustement", delta, "Manuel"
    Call RefreshStocks: Call MarkStale
End Sub

Public Sub LogMovement(ByVal product As String, ByVal kind As String, ByVal qty As Long, ByVal ref As String)
    AppendRow "tblT_Mouvements", Array(NextId("tblT_Mouvements", "M"), CDbl(Now), product, kind, qty, ref)
End Sub

Public Sub RefreshStocks()
    Dim lo As ListObject: Set lo = Tbl("tblT_Stock")
    If Not lo.DataBodyRange Is Nothing Then lo.ListColumns("Etat").DataBodyRange.FormulaR1C1 = "=IF(RC[-2]<RC[-1],""ALERTE"",""OK"")"
    ThisWorkbook.Worksheets("Menu").Range("B8").Value2 = Tbl("tblT_Produits").ListRows.count
    ThisWorkbook.Worksheets("Menu").Range("C8").Value2 = Tbl("tblT_Fournisseurs").ListRows.count
    UpdateOperationalKPI
End Sub

Public Sub Calibrate(Optional ByVal silent As Boolean = False)
    On Error GoTo Failed
    Dim a As Variant, dict As Object, i As Long, key As String, v As Variant, n As Long
    Dim total As Double, squares As Double, mean As Double, sigma As Double
    Dim margin As Double, marginN As Long, delay As Double, delayN As Long, late As Long, out() As Variant
    Dim suppliers As Object, frequencies As Object, supplierData As Variant, supplierKey As String
    Dim expectedDelay As Double, sourceRows As Long, maxCount As Long, modalCost As Double, product As String
    product = CStr(Param("ProduitSimulation"))
    Require FindRow("tblT_Produits", product) > 0, "Choisissez un produit existant."
    Set suppliers = CreateObject("Scripting.Dictionary")
    Set frequencies = CreateObject("Scripting.Dictionary")
    supplierData = Tbl("tblT_Fournisseurs").DataBodyRange.Value2
    For i = 1 To UBound(supplierData, 1)
        If CStr(supplierData(i, 1)) <> "" Then suppliers.Add CStr(supplierData(i, 1)), Array(CDbl(supplierData(i, 3)), CDbl(supplierData(i, 4)))
    Next i
    Set dict = CreateObject("Scripting.Dictionary")
    a = Tbl("tblT_Commandes").DataBodyRange.Value2
    For i = 1 To UBound(a, 1)
        If CStr(a(i, 4)) = product And CStr(a(i, 8)) <> "Annulee" Then
            key = CStr(Fix(CDbl(a(i, 2))))
            If Not dict.Exists(key) Then dict.Add key, 0#
            dict(key) = dict(key) + CDbl(a(i, 5))
            If a(i, 9) = "DataCo" Then
                margin = margin + CDbl(a(i, 15)) / CDbl(a(i, 5)): marginN = marginN + 1
                delay = delay + CDbl(a(i, 12)): delayN = delayN + 1
                If a(i, 13) = "Late delivery" Then late = late + 1
                supplierKey = CStr(a(i, 10))
                If suppliers.Exists(supplierKey) Then
                    v = suppliers(supplierKey)
                    expectedDelay = expectedDelay + CDbl(v(0)): sourceRows = sourceRows + 1
                    key = CStr(v(1))
                    If Not frequencies.Exists(key) Then frequencies.Add key, 0
                    frequencies(key) = frequencies(key) + 1
                End If
            End If
        End If
    Next i
    n = dict.count: Require n >= 2, "Au moins deux jours de demande sont necessaires."
    For Each v In dict.Items: total = total + v: Next v
    mean = total / n
    For Each v In dict.Items: squares = squares + (v - mean) ^ 2: Next v
    sigma = Sqr(squares / (n - 1))
    ReDim out(1 To n, 1 To 2): i = 0
    For Each v In dict.Keys: i = i + 1: out(i, 1) = CDbl(v): out(i, 2) = dict(v): Next v
    WriteTable "tblDemande_Journaliere", out, n, 2
    SetParam "mu", mean: SetParam "sigma", sigma
    If marginN > 0 Then SetParam "Crup", margin / marginN
    If sourceRows > 0 Then
        delay = Fix(expectedDelay / sourceRows + 0.5): If delay < 1 Then delay = 1
        SetParam "L", delay
        For Each v In frequencies.Keys
            If frequencies(v) > maxCount Then maxCount = frequencies(v): modalCost = CDbl(v)
        Next v
        SetParam "Ccmd", modalCost
    End If
    SuggestProductPolicies
    MarkStale
    If Not silent Then MsgBox "Calibration " & product & " : " & n & " jours, moyenne " & Format$(mean, "0.00") & ", sigma " & Format$(sigma, "0.00"), vbInformation
    Exit Sub
Failed:
    If silent Then Err.Raise Err.Number, Err.source, Err.Description Else MsgBox Err.Description, vbExclamation
End Sub

Public Sub SuggestProductPolicies()
    ' Daily review: protect L+1 days, since ordering follows the day's sales.
    ' These are starting policies to compare, not a guarantee of service or optimality.
    Dim mean As Double, sigma As Double, lead As Double, holding As Double, fee As Double
    Dim target As Double, z As Double, safety As Double, demand As Double, lot As Double
    Dim policies(1 To 3, 1 To 3) As Variant, i As Long, baseS As Long, baseQ As Long, initial As Long
    Dim events As Boolean, message As String, code As Long
    On Error GoTo Failed
    events = Application.EnableEvents
    mean = NumberValue(Param("mu"), "mu", 0)
    sigma = NumberValue(Param("sigma"), "sigma", 0)
    lead = NumberValue(Param("L"), "L", 1, True)
    holding = NumberValue(Param("Cpos"), "Cpos", 0)
    fee = NumberValue(Param("Ccmd"), "Ccmd", 0)
    target = NumberValue(Param("SeuilService"), "SeuilService", 0)
    Require target <= 1, "SeuilService entre 0 et 1."
    ' Bound the normal quantile, including service targets of zero and one.
    target = Application.Max(0.5, Application.Min(0.9999, target))
    z = Application.WorksheetFunction.Norm_S_Inv(target)
    demand = mean * (lead + 1): safety = sigma * Sqr(lead + 1)
    ' Economic lot uses daily holding cost. When costs are zero, use a lead-time lot.
    If holding > 0 And fee > 0 Then
        lot = Sqr(2# * mean * fee / holding)
    Else
        lot = mean * lead
    End If
    lot = Application.Max(1, mean, lot)
    baseS = PolicyUnits(demand + z * safety, 0)
    baseQ = PolicyUnits(lot, 1)
    initial = PolicyUnits(baseS + baseQ / 2#, 0)
    For i = 1 To 3: policies(i, 1) = initial: Next i
    policies(1, 2) = PolicyUnits(demand + (z + 0.75) * safety, 0)
    policies(1, 3) = PolicyUnits(lot * 1.35, 1)
    policies(2, 2) = PolicyUnits(demand + Application.Max(0, z - 0.75) * safety, 0)
    policies(2, 3) = PolicyUnits(Application.Max(mean, lot * 0.7), 1)
    policies(3, 2) = baseS: policies(3, 3) = baseQ
    Application.EnableEvents = False
    Tbl("tblScenarios").DataBodyRange.columns(3).Resize(3, 3).Value2 = policies
    SetParam "S0", initial: SetParam "s", baseS: SetParam "Q", baseQ
    MarkStale
    Application.EnableEvents = events
    Exit Sub
Failed:
    code = Err.Number: message = Err.Description
    Application.EnableEvents = events
    Err.Raise code, "SuggestProductPolicies", message
End Sub

Private Function PolicyUnits(ByVal value As Double, ByVal minimum As Long) As Long
    Require value <= 1000000, "Suggestion superieure a 1000000 unites : ajustez les parametres."
    PolicyUnits = CLng(Application.Max(minimum, -Int(-value)))
End Function

Public Sub SyncScenarioInitialStock()
    ' All policies use the same initial stock for a fair paired comparison.
    Tbl("tblScenarios").DataBodyRange.columns(3).Value2 = Param("S0")
End Sub



