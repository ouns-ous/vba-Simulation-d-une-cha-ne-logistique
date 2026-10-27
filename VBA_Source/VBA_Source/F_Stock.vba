Option Explicit
Private Sub UserForm_Initialize()
FillProducts cboProduct
txtDelta.value = "0"
txtThreshold.value = "260"
RefreshList
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSearch_Click()
RefreshList
End Sub

Private Sub btnNew_Click()
txtDelta.value = "0"
End Sub

Private Sub lstRecords_Click()
On Error GoTo Failed
If lstRecords.ListIndex < 0 Then Exit Sub
cboProduct.value = lstRecords.list(lstRecords.ListIndex, 0)
txtThreshold.value = lstRecords.list(lstRecords.ListIndex, 2)
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnSave_Click()
On Error GoTo Failed
Require IsNumeric(txtDelta.value), "Mouvement numerique obligatoire."
Require CDbl(txtDelta.value) = Fix(CDbl(txtDelta.value)), "Mouvement entier obligatoire."
AdjustStock CStr(cboProduct.value), CLng(txtDelta.value), CLng(NumberValue(txtThreshold.value, "Seuil", 0, True))
txtDelta.value = "0"
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnTable_Click()
OpenApplication 5
End Sub

Private Sub RefreshList()
FillRecords lstRecords, "tblT_Stock", Array(1, 2, 3, 4), CStr(txtSearch.value)
End Sub
Private Sub cboProduct_Change()
Dim i As Long
i = FindRow("tblT_Stock", CStr(cboProduct.value))
If i > 0 Then txtThreshold.value = Tbl("tblT_Stock").DataBodyRange.Cells(i, 3).Value2
End Sub

