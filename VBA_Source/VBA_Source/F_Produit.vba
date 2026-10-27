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
txtKey.Locked = False: txtKey.value = "": txtName.value = "": txtPrice.value = ""
End Sub

Private Sub lstRecords_Click()
On Error GoTo Failed
If lstRecords.ListIndex < 0 Then Exit Sub
Dim i As Long
i = FindRow("tblT_Produits", CStr(lstRecords.list(lstRecords.ListIndex, 0)))
With Tbl("tblT_Produits").DataBodyRange
txtKey.value = .Cells(i, 1).Value2: txtName.value = .Cells(i, 2).Value2: txtPrice.value = .Cells(i, 3).Value2
End With
txtKey.Locked = True
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnSave_Click()
On Error GoTo Failed
SaveProduct CStr(txtKey.value), CStr(txtName.value), NumberValue(txtPrice.value, "Prix")
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnDelete_Click()
On Error GoTo Failed
DeleteProduct CStr(txtKey.value)
txtKey.value = "": txtName.value = "": txtPrice.value = "": txtKey.Locked = False
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub RefreshList()
FillRecords lstRecords, "tblT_Produits", Array(1, 2, 3), CStr(txtSearch.value)
End Sub
