Option Explicit
Private Sub UserForm_Initialize()
RefreshList
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSearch_Click()
RefreshList
End Sub

Private Sub btnNew_Click()
txtKey.Locked = False: txtKey.value = "": txtName.value = "": txtDelay.value = "": txtCost.value = ""
End Sub

Private Sub lstRecords_Click()
On Error GoTo Failed
If lstRecords.ListIndex < 0 Then Exit Sub
Dim i As Long
i = FindRow("tblT_Fournisseurs", CStr(lstRecords.list(lstRecords.ListIndex, 0)))
With Tbl("tblT_Fournisseurs").DataBodyRange
txtKey.value = .Cells(i, 1).Value2: txtName.value = .Cells(i, 2).Value2
txtDelay.value = .Cells(i, 3).Value2: txtCost.value = .Cells(i, 4).Value2
End With
txtKey.Locked = True
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnSave_Click()
On Error GoTo Failed
SaveSupplier CStr(txtKey.value), CStr(txtName.value), NumberValue(txtDelay.value, "Delai"), NumberValue(txtCost.value, "Cout")
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnDelete_Click()
On Error GoTo Failed
DeleteSupplier CStr(txtKey.value)
txtKey.Locked = False: txtKey.value = "": txtName.value = ""
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub RefreshList()
FillRecords lstRecords, "tblT_Fournisseurs", Array(1, 2, 3, 4), CStr(txtSearch.value)
End Sub
