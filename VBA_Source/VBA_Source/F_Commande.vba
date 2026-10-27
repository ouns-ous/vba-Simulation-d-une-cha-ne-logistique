Option Explicit
Private Sub UserForm_Initialize()
FillProducts cboProduct
txtDate.value = Format$(Date, "dd/mm/yyyy")
txtClient.value = "Client 001"
txtKey.Locked = True
RefreshList
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSearch_Click()
RefreshList
End Sub

Private Sub btnNew_Click()
txtKey.value = "": txtQty.value = "": txtDate.value = Format$(Date, "dd/mm/yyyy")
End Sub

Private Sub lstRecords_Click()
On Error GoTo Failed
If lstRecords.ListIndex < 0 Then Exit Sub
txtKey.value = lstRecords.list(lstRecords.ListIndex, 0)
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnSave_Click()
On Error GoTo Failed
Dim key As String
key = CreateOrder(CStr(cboProduct.value), CLng(NumberValue(txtQty.value, "Quantite", 1, True)), ParseDate(CStr(txtDate.value)), CStr(txtClient.value))
txtKey.value = key
MsgBox "Commande " & key & " enregistree. Les unites indisponibles sont des ventes perdues.", vbInformation
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnCancel_Click()
On Error GoTo Failed
CancelOrder CStr(txtKey.value)
MsgBox "Commande annulee. Stock vendu restitue.", vbInformation
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub RefreshList()
FillRecords lstRecords, "tblT_Commandes", Array(1, 2, 4, 5, 8), CStr(txtSearch.value)
End Sub
