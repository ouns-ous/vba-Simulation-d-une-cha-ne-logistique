Attribute VB_Name = "TestSuggestions"
Option Explicit

Public Function VerifyProductSuggestions() As String
    Dim saved As Variant, scenarios As Variant, sh As Worksheet, events As Boolean
    Dim f As Object, selector As Object, result As String, expected As Variant, i As Long
    On Error GoTo Failed
    Set sh = ThisWorkbook.Worksheets("Feuille_Controle")
    saved = sh.Range("C6:C21").Value2
    scenarios = Tbl("tblScenarios").DataBodyRange.Value2
    events = Application.EnableEvents
    Application.EnableEvents = False
    Set f = F_Main
    Set selector = f.tabsMain.Pages(6).Controls("cboSimulationProduct")
    For i = 1 To 4
        selector.Value = "P00" & i
        Require Param("ProduitSimulation") = "P00" & i, "Selecteur produit non applique."
        Require f.tabsMain.Pages(6).Controls("par0").Value = CStr(Param("S0")), "S0 du formulaire obsolete."
        Require f.tabsMain.Pages(7).Controls("scenarios3").Value = CStr(Param("s")), "Suggestion s obsolete."
        Require f.tabsMain.Pages(7).Controls("scenarioq3").Value = CStr(Param("Q")), "Suggestion Q obsolete."
    Next i
    ' Reloading or saving manual policies must not silently replace them.
    f.tabsMain.Pages(6).Controls("par0").Value = "777"
    f.tabsMain.Pages(7).Controls("scenarios1").Value = "123"
    f.tabsMain.Pages(7).Controls("scenarioq1").Value = "456"
    Require f.ApplySettings(False, False), "Enregistrement manuel refuse."
    f.LoadSettings
    Require Param("S0") = 777, "S0 manuel ecrase."
    Require Tbl("tblScenarios").DataBodyRange.Cells(1, 4).Value2 = 123, "s manuel ecrase."
    Require Tbl("tblScenarios").DataBodyRange.Cells(1, 5).Value2 = 456, "Q manuel ecrase."
    For i = 1 To 3
        Require Tbl("tblScenarios").DataBodyRange.Cells(i, 3).Value2 = 777, "S0 scenario desynchronise."
    Next i
    Calibrate True
    Require Param("S0") <> 777, "Recalibration sans nouvelles suggestions."
    Application.EnableEvents = True
    sh.Range("C21").Value2 = "P001"
    Require Param("S0") = 471, "Changement produit feuille non recalcule."
    Require ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2 = "PERIME", "Resultats non invalides."
    sh.Range("C6").Value2 = 888
    Require Tbl("tblScenarios").DataBodyRange.Cells(3, 3).Value2 = 888, "S0 feuille non synchronise."
    Application.EnableEvents = False
    SetParam "mu", 0: SetParam "sigma", 0: SetParam "Cpos", 0: SetParam "Ccmd", 0
    SetParam "SeuilService", 0
    SuggestProductPolicies
    Require Param("s") = 0 And Param("Q") >= 1, "Demande et couts nuls invalides."
    SetParam "SeuilService", 1
    SuggestProductPolicies
    SetParam "mu", 1000000: SetParam "L", 3650
    On Error Resume Next
    SuggestProductPolicies
    expected = Err.Number: Err.Clear
    On Error GoTo Failed
    Require expected <> 0, "Depassement de capacite non refuse."
    Require Application.EnableEvents = False, "Etat des evenements altere apres erreur."
    result = "PASS: 4 selecteurs, rechargement, valeurs manuelles, recalibration, evenements feuille, S0 commun, bornes 0/1, couts nuls et capacite."
    GoTo Cleanup
Failed:
    result = "FAIL: " & Err.Number & " " & Err.Description
Cleanup:
    Application.EnableEvents = False
    sh.Range("C6:C21").Value2 = saved
    Tbl("tblScenarios").DataBodyRange.Value2 = scenarios
    If Not f Is Nothing Then Unload f
    Application.EnableEvents = events
    VerifyProductSuggestions = result
End Function
