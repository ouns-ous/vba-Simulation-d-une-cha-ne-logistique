Option Explicit

Private Sub Workbook_Open()
    EnsureWorkbookVisible
    ProtectData
    RefreshStocks
    ThisWorkbook.Worksheets("Menu").Activate
    OpenMain
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    If sh.name = "Feuille_Controle" Then
        If Not Intersect(target, sh.Range("C6:C21")) Is Nothing Then MarkStale
    ElseIf sh.name = "Scenarios" Then
        If Not Intersect(target, sh.ListObjects("tblScenarios").Range) Is Nothing Then MarkStale
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim form As Object
    For Each form In VBA.UserForms
        If TypeName(form) = "F_Main" Then Unload form: Exit For
    Next form
End Sub
