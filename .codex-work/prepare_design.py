from pathlib import Path
import re
root = Path(__file__).resolve().parent.parent
src = root / 'VBA_Source'
def native(name):
    return (root / '.codex-work' / f'design-before-{name}.vba').read_bytes().decode('utf-8').replace('\r\r\n','\n').replace('\r\n','\n')
def proc(text, name, replacement):
    pattern = rf'(?m)^(?:Private|Public) Sub {name}\([^\n]*\).*?^End Sub'
    text, count = re.subn(pattern, lambda m: replacement.strip(), text, flags=re.S)
    assert count == 1, (name,count)
    return text
layout = native('modFormLayout')
for old,new in {
 'RGB(12, 35, 64)':'RGB(39, 49, 63)',
 'RGB(92, 114, 141)':'RGB(110, 121, 139)',
 'RGB(31, 100, 201)':'RGB(85, 101, 160)',
 'RGB(232, 240, 250)':'RGB(243, 245, 250)',
 'RGB(6, 30, 58)':'RGB(255, 255, 255)',
 'RGB(47, 128, 237)':'RGB(85, 101, 160)',
 'RGB(248, 251, 255)':'RGB(255, 255, 255)',
 'RGB(250, 231, 220)':'RGB(247, 248, 250)',
 'RGB(246, 249, 253)':'RGB(247, 248, 250)',
 'RGB(226, 236, 248)':'RGB(255, 255, 255)',
 'RGB(195, 205, 221)':'RGB(205, 211, 222)',
}.items(): layout=layout.replace(old,new)
layout=layout.replace('p.Controls(name).ForeColor = Blue()', 'p.Controls(name).ForeColor = Ink()')
layout=layout.replace('TextStyle c, 10.5\n                c.BackStyle', 'TextStyle c, 10.5, True\n                c.BackStyle')
layout=layout.replace('TextStyle c, 11\n                c.BackColor', 'TextStyle c, 11, True\n                c.BackColor')
layout=layout.replace('w = p.InsideWidth: h = p.InsideHeight', 'w = mPageWidth: h = mPageHeight')
layout=layout.replace('Option Explicit','Option Explicit\nPrivate mPageWidth As Single\nPrivate mPageHeight As Single',1)
layout=layout.replace('p.Controls("lblParam15").Caption = "Produit historique"','p.Controls("lblParam15").Caption = "Produit historique"')
layout=layout.replace('    p.Controls("btnEvolution").BackColor = RGB(229, 240, 253)','    p.Controls("btnEvolution").BackColor = RGB(247, 248, 250)')
layout=layout.replace('    p.Controls("btnEvolution").ForeColor = RGB(25, 83, 159)','    p.Controls("btnEvolution").ForeColor = Ink()')
layout=layout.replace('    p.Controls("btnResultsRefresh").BackColor = RGB(228, 245, 234)','    p.Controls("btnResultsRefresh").BackColor = RGB(247, 248, 250)')
layout=layout.replace('    p.Controls("btnResultsRefresh").ForeColor = RGB(31, 128, 78)','    p.Controls("btnResultsRefresh").ForeColor = Ink()')
layout=proc(layout,'HomePage','''Private Sub HomePage(ByVal p As Object)
    Dim w As Single, h As Single, cw As Single, x As Single, i As Long, c As Object, titles As Variant, values As Variant
    Dim sideW As Single, mainX As Single, tileW As Single, tileH As Single
    w = mPageWidth: h = Application.Max(440, mPageHeight)
    titles = Array("Commandes historiques", "Produits", "Fournisseurs", "Produits sous seuil", "Appro. en attente")
    values = Array(Tbl("tblT_Commandes").ListRows.Count, Tbl("tblT_Produits").ListRows.Count, Tbl("tblT_Fournisseurs").ListRows.Count, ThisWorkbook.Worksheets("Menu").Range("E35").Value2, ThisWorkbook.Worksheets("Menu").Range("E36").Value2)
    cw = (w - 48 - 4 * 14) / 5
    For i = 0 To 4
        x = 18 + i * (cw + 14)
        Backdrop p, "homeCard" & i, x, 14, cw, 104, RGB(255, 255, 255)
        Set c = Heading(p, "homeMetricTitle" & i)
        c.Caption = titles(i): Bounds c, x + 6, 27, cw - 12, 28: TextStyle c, 11, True: c.TextAlign = 2
        Set c = Heading(p, "homeMetricValue" & i)
        c.Caption = Format$(values(i), "#,##0"): Bounds c, x + 6, 65, cw - 12, 37: TextStyle c, 27, True: c.TextAlign = 2
    Next i
    sideW = 214: mainX = sideW + 36
    Backdrop p, "homeSidebar", 18, 136, sideW, h - 154, RGB(255, 255, 255)
    Set c = Heading(p, "homeManagement")
    c.Caption = "Gestion": Bounds c, 30, 148, sideW - 24, 22: TextStyle c, 13, True
    For i = 1 To 5
        Place p, "nav" & i, 30, 182 + (i - 1) * 43, sideW - 24, 35
        ButtonStyle p.Controls("nav" & i)
    Next i
    p.Controls("nav3").Caption = "Commandes clients"
    p.Controls("nav4").Caption = "Approvisionnements"
    p.Controls("nav5").Caption = "Stocks et seuils"
    Place p, "lblWelcome", mainX, 136, w - mainX - 18, 31
    p.Controls("lblWelcome").Caption = "Gestion et analyse logistique": TextStyle p.Controls("lblWelcome"), 19, True
    Place p, "lblIntro", mainX, 176, w - mainX - 18, 32
    p.Controls("lblIntro").Caption = "Choisissez une rubrique pour ouvrir les outils et les rapports."
    TextStyle p.Controls("lblIntro"), 11, True: p.Controls("lblIntro").ForeColor = Muted()
    tileW = (w - mainX - 32) / 2: tileH = Application.Min(79, (h - 260) / 2)
    For i = 6 To 9
        x = mainX + ((i - 6) Mod 2) * (tileW + 14)
        Place p, "nav" & i, x, 224 + ((i - 6) \\ 2) * (tileH + 16), tileW, tileH
        ButtonStyle p.Controls("nav" & i): p.Controls("nav" & i).Font.Size = 14
    Next i
    p.Controls("nav9").Caption = "Exports et aide"
    p.Controls("lblCounters").Visible = False
    Place p, "lblHomeState", mainX, h - 28, w - mainX - 18, 20
    p.Controls("lblHomeState").ForeColor = Muted()
End Sub''')
layout=proc(layout,'LayoutDataCoForm','''Public Sub LayoutDataCoForm(ByVal view As Object)
    Dim tabs As Object, p As Object, i As Long, active As Long, w As Single, h As Single, c As Object, nw As Single
    Dim stage As String
    On Error GoTo Failed
    stage = "Cadre"
    view.BackColor = Canvas()
    view.ScrollBars = 0: view.ScrollTop = 0: view.ScrollLeft = 0
    w = view.InsideWidth: h = view.InsideHeight
    Set tabs = view.Controls("tabsMain")
    active = tabs.Value
    Backdrop view, "lblBrandBar", 0, 0, w, 80, RGB(255, 255, 255)
    Backdrop view, "lblAccentBar", 0, 79, w, 2, Accent()
    Backdrop view, "lblFooterBar", 0, h - 47, w, 47, RGB(255, 255, 255)
    Place view, "lblAppTitle", 22, 12, w - 44, 34
    TextStyle view.Controls("lblAppTitle"), 25, True
    view.Controls("lblAppTitle").Caption = "DataCo Global"
    Place view, "lblAppSubtitle", 24, 51, w - 48, 20
    TextStyle view.Controls("lblAppSubtitle"), 11, True
    view.Controls("lblAppSubtitle").ForeColor = Muted()
    nw = (w - 44 - 10 * 5) / 11
    For i = 0 To 10
        Set c = view.Controls("designNav" & i)
        Bounds c, 22 + i * (nw + 5), 90, nw, 34
        ButtonStyle c: c.Font.Size = 10.5: c.Font.Italic = True
    Next i
    tabs.Style = 2: tabs.BackColor = Canvas()
    Bounds tabs, 12, 133, w - 24, h - 185
    mPageWidth = Application.Max(920, tabs.Width - 6)
    mPageHeight = Application.Max(408, tabs.Height - 6)
    For i = 0 To tabs.Pages.Count - 1
        stage = "Page " & i
        tabs.Value = i
        Set p = tabs.Pages(i)
        p.Zoom = 100
        p.ScrollBars = 0: p.KeepScrollBarsVisible = 0
        p.ScrollLeft = 0: p.ScrollTop = 0
        StylePage p
        Backdrop p, "designPageBackground", 0, 0, mPageWidth, mPageHeight, Canvas()
        Select Case i
            Case 0: HomePage p
            Case 1: CatalogPage p, "pr_", Array("txtKey", "txtName", "txtPrice"), Array("Reference", "Designation", "Prix ($)")
            Case 2: CatalogPage p, "fu_", Array("txtKey", "txtName", "txtDelay", "txtCost"), Array("Code", "Fournisseur", "Delai (j)", "Cout ($)")
            Case 3: CatalogPage p, "co_", Array("txtKey", "cboProduct", "txtQty", "txtDate", "txtClient"), Array("Commande", "Date", "Produit", "Quantite", "Statut")
            Case 4: CatalogPage p, "ap_", Array("txtKey", "cboSupplier", "cboProduct", "txtQty", "txtDate", "txtReceipt"), Array("Commande", "Fournisseur", "Produit", "Quantite", "Statut")
            Case 5: CatalogPage p, "st_", Array("cboProduct", "txtDelta", "txtThreshold"), Array("Produit", "Stock", "Seuil", "Etat")
            Case 6: SettingsPage p
            Case 7: SimulationPage p
            Case 8: ResultsPage p
            Case 9: ExportPage p
            Case 10: ReportPage p
        End Select
        p.Controls("designPageBackground").ZOrder 1
        p.ScrollWidth = mPageWidth: p.ScrollHeight = mPageHeight
        If i = 0 Then p.ScrollHeight = Application.Max(440, mPageHeight)
    Next i
    tabs.Value = active
    stage = "Pied de page"
    Place view, "designZoomOut", 22, h - 37, 38, 30
    Place view, "designZoomReset", 70, h - 37, 61, 30
    Place view, "designZoomIn", 141, h - 37, 38, 30
    Place view, "designZoomFit", 189, h - 37, 98, 30
    For Each c In view.Controls
        If Left$(c.Name, 10) = "designZoom" Then ButtonStyle c
    Next c
    Place view, "lblStatus", 306, h - 31, Application.Max(80, w - 574), 23
    TextStyle view.Controls("lblStatus"), 10.5, True
    view.Controls("lblStatus").ForeColor = Muted()
    Place view, "btnHome", w - 252, h - 37, 112, 30
    Place view, "btnClose", w - 128, h - 37, 112, 30
    ButtonStyle view.Controls("btnHome"): ButtonStyle view.Controls("btnClose")
    Exit Sub
Failed:
    Err.Raise Err.Number, "LayoutDataCoForm", stage & " : " & Err.Description
End Sub

Public Sub PaintDataCoNavigation(ByVal view As Object)
    Dim i As Long, c As Object
    For i = 0 To 10
        Set c = view.Controls("designNav" & i)
        c.BackColor = RGB(255, 255, 255): c.ForeColor = Ink()
        If i = view.Controls("tabsMain").Value Then
            c.BackColor = RGB(228, 232, 243): c.ForeColor = Blue()
        End If
    Next i
End Sub

Public Sub ZoomDataCoPages(ByVal view As Object, ByVal percent As Long, Optional ByVal pageIndex As Long = -1)
    Dim p As Object
    If pageIndex < 0 Then pageIndex = view.Controls("tabsMain").Value
    Set p = view.Controls("tabsMain").Pages(pageIndex)
    p.Zoom = percent
    p.ScrollBars = 3
    p.KeepScrollBarsVisible = 0
    p.ScrollTop = 0: p.ScrollLeft = 0
    view.Controls("designZoomReset").Caption = CStr(percent) & "%"
End Sub

Public Function FitDataCoZoom(ByVal view As Object) As Long
    Dim tabs As Object
    Set tabs = view.Controls("tabsMain")
    FitDataCoZoom = Application.Max(50, Application.Min(100, Int(Application.Min((tabs.Width - 12) / mPageWidth, (tabs.Height - 12) / tabs.SelectedItem.ScrollHeight) * 100)))
End Function''')
form=native('F_Main')
form=form.replace('Private mReportHeaders As Variant','Private mReportHeaders As Variant\nPrivate mDesignButtons As Collection\nPrivate mContentZoom As Long\nPrivate mPageZoom(0 To 10) As Long',1)
form=form.replace('    InstallReportViewer\n','    InstallReportViewer\n    InstallDesignControls\n',1)
form=form.replace('    LayoutDataCoForm Me\n','    LayoutDataCoForm Me\n    ApplyDesignZoom\n    PaintDataCoNavigation Me\n',1)
form=form.replace('    lblStatus.Caption = tabsMain.SelectedItem.Caption','    PaintDataCoNavigation Me\n    mContentZoom = mPageZoom(tabsMain.Value)\n    ZoomDataCoPages Me, mContentZoom\n    lblStatus.Caption = tabsMain.SelectedItem.Caption',1)
form=form.replace('    If CloseMode = 0 Or CloseMode = 1 Then ShowExcelHome','    Set mDesignButtons = Nothing\n    If CloseMode = 0 Or CloseMode = 1 Then ShowExcelHome',1)
form=form.replace('    Me.Height = Application.Min(760, Application.Height - 24)','    Me.Height = Application.Min(760, Application.Height - 24)',1)
# Repaint only the dashboard figures after the existing KPI refresh.
form=form.replace('    lblHomeState.Caption = "Resultats : " & ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2','    lblHomeState.Caption = "Resultats : " & ThisWorkbook.Worksheets("Export_Decision").Range("E1").Value2\n    If mReady And mContentZoom > 0 And Not mLayoutBusy Then ArrangeInterface',1)
form += '''

Private Sub InstallDesignControls()
    Dim i As Long, c As MSForms.CommandButton, handler As CDataCoDesignButton
    Dim names As Variant, actions As Variant, captions As Variant
    Set mDesignButtons = New Collection
    For i = 0 To tabsMain.Pages.Count - 1
        mPageZoom(i) = 80
        Set c = Me.Controls.Add("Forms.CommandButton.1", "designNav" & i, True)
        c.Caption = tabsMain.Pages(i).Caption: c.Tag = "page:" & i
        Set handler = New CDataCoDesignButton
        Set handler.Button = c: Set handler.Owner = Me
        mDesignButtons.Add handler
    Next i
    mPageZoom(0) = 100
    names = Array("designZoomOut", "designZoomReset", "designZoomIn", "designZoomFit")
    actions = Array("zoom:-10", "zoom:100", "zoom:10", "fit")
    captions = Array("-", "100%", "+", "Ajuster")
    For i = 0 To 3
        Set c = Me.Controls.Add("Forms.CommandButton.1", CStr(names(i)), True)
        c.Caption = captions(i): c.Tag = actions(i)
        Set handler = New CDataCoDesignButton
        Set handler.Button = c: Set handler.Owner = Me
        mDesignButtons.Add handler
    Next i
End Sub

Public Sub HandleDesignAction(ByVal action As String)
    If Left$(action, 5) = "page:" Then
        SelectPage CLng(Mid$(action, 6))
    Else
        Select Case action
            Case "fit": mContentZoom = FitDataCoZoom(Me)
            Case "zoom:100": mContentZoom = 100
            Case "zoom:-10": mContentZoom = Application.Max(50, mContentZoom - 10)
            Case "zoom:10": mContentZoom = Application.Min(150, mContentZoom + 10)
        End Select
        mLayoutBusy = True
        mPageZoom(tabsMain.Value) = mContentZoom
        ZoomDataCoPages Me, mContentZoom
        mLayoutBusy = False
    End If
End Sub

Private Sub ApplyDesignZoom()
    Dim i As Long
    For i = 0 To 10
        ZoomDataCoPages Me, mPageZoom(i), i
    Next i
    mContentZoom = mPageZoom(tabsMain.Value)
    Me.Controls("designZoomReset").Caption = CStr(mContentZoom) & "%"
End Sub
'''
classcode='''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CDataCoDesignButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Button As MSForms.CommandButton
Public Owner As Object

Private Sub Button_Click()
    Owner.HandleDesignAction CStr(Button.Tag)
End Sub
'''
(src/'modFormLayout.bas').write_text('Attribute VB_Name = "modFormLayout"\n'+layout, encoding='utf-8',newline='\n')
# Keep the current exported form designer header; only its code changes.
old=(src/'F_Main.frm').read_text(encoding='utf-8')
(src/'F_Main.frm').write_text(old[:old.index('Option Explicit')]+form,encoding='utf-8',newline='\n')
(src/'CDataCoDesignButton.cls').write_text(classcode,encoding='utf-8',newline='\n')
print('Prepared layout, form presentation controls and button event class')
