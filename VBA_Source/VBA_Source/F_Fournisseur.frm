VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Fournisseur 
   Caption         =   "Gestion des fournisseurs"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   OleObjectBlob   =   "F_Fournisseur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Fournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

