VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Appro 
   Caption         =   "Approvisionnements et receptions"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   OleObjectBlob   =   "F_Appro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Appro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
FillProducts cboProduct
FillSuppliers cboSupplier
txtDate.value = Format$(Date, "dd/mm/yyyy")
txtReceipt.value = Format$(Date, "dd/mm/yyyy")
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
txtKey.value = "": txtQty.value = ""
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
txtKey.value = CreateSupply(CStr(cboSupplier.value), CStr(cboProduct.value), CLng(NumberValue(txtQty.value, "Quantite", 1, True)), ParseDate(CStr(txtDate.value)))
MsgBox "Approvisionnement cree.", vbInformation
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnReceive_Click()
On Error GoTo Failed
ReceiveSupply CStr(txtKey.value), ParseDate(CStr(txtReceipt.value))
MsgBox "Reception enregistree et stock actualise.", vbInformation
RefreshList
Exit Sub
Failed:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub RefreshList()
FillRecords lstRecords, "tblT_Approvisionnements", Array(1, 3, 4, 5, 8), CStr(txtSearch.value)
End Sub

