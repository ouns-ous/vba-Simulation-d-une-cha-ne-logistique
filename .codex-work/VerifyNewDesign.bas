Option Explicit
Public Function VerifyNewDesign() As String
    Dim i As Long, p As Object, c As Object, s As String, z As Long
    Load F_Main
    If Len(F_Main.DisplayError) > 0 Then Err.Raise 513, , F_Main.DisplayError
    s = "Layout: OK" & vbCrLf
    For i = 0 To 10
        F_Main.HandleDesignAction "page:" & i
        If F_Main.Controls("tabsMain").Value <> i Then Err.Raise 514, , "Navigation failed " & i
        Set p = F_Main.Controls("tabsMain").Pages(i)
        If p.Controls("designPageBackground").Width < 900 Then Err.Raise 515, , "Page dimensions " & i
        s = s & "Page " & i & ": " & p.Caption & " OK" & vbCrLf
    Next i
    Set p = F_Main.Controls("tabsMain").Pages(6)
    Set c = p.Controls("cboSimulationProduct")
    If Abs(c.Left - p.Controls("par15").Left) > 0.1 Or Abs(c.Top - p.Controls("par15").Top) > 0.1 Then Err.Raise 516, , "Product position"
    If c.ListCount <> 4 Then Err.Raise 517, , "Product choices"
    F_Main.HandleDesignAction "page:1"
    F_Main.HandleDesignAction "zoom:100"
    F_Main.HandleDesignAction "zoom:10"
    If F_Main.Controls("tabsMain").Pages(1).Zoom <> 110 Then Err.Raise 518, , "Zoom in"
    F_Main.HandleDesignAction "zoom:-10"
    If F_Main.Controls("tabsMain").Pages(1).Zoom <> 100 Then Err.Raise 519, , "Zoom out"
    F_Main.HandleDesignAction "fit"
    z = F_Main.Controls("tabsMain").Pages(1).Zoom
    If z < 50 Or z > 100 Then Err.Raise 520, , "Fit zoom"
    s = s & "Product: 4 choices, aligned" & vbCrLf & "Zoom +/-/100%/Fit: OK (" & z & "%)" & vbCrLf
    F_Main.HandleDesignAction "page:0"
    Set p = F_Main.Controls("tabsMain").Pages(0)
    For i = 0 To 4
        s = s & p.Controls("homeMetricTitle" & i).Caption & ": " & p.Controls("homeMetricValue" & i).Caption & vbCrLf
    Next i
    VerifyNewDesign = s
    Unload F_Main
End Function

Public Sub ShowNewDesign()
    F_Main.SelectPage 0
    F_Main.Show vbModeless
End Sub
