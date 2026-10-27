$ErrorActionPreference = 'Stop'
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $xl.Workbooks.Item('DataCo_Logistique.xlsm')
$helper = $wb.VBProject.VBComponents.Add(1)
$helper.Name = 'TempVerifySelectorPosition'
$code = @'
Option Explicit
Public Function VerifyProductSelectorPosition() As String
    Dim p As Object, c As Object, anchor As Object, aligned As Boolean
    Set p = F_Main.Controls("tabsMain").Pages(6)
    Set c = p.Controls("cboSimulationProduct")
    Set anchor = p.Controls("par15")
    aligned = Abs(c.Left - anchor.Left) < 0.1 And Abs(c.Top - anchor.Top) < 0.1 And Abs(c.Width - anchor.Width) < 0.1 And Abs(c.Height - anchor.Height) < 0.1
    VerifyProductSelectorPosition = "Aligned=" & CStr(aligned) & "; Left=" & CStr(c.Left) & "; Top=" & CStr(c.Top) & "; Width=" & CStr(c.Width) & "; Height=" & CStr(c.Height) & "; Choices=" & CStr(c.ListCount)
    Unload F_Main
End Function
'@
try {
    $helper.CodeModule.AddFromString($code)
    $result = $xl.Run("'DataCo_Logistique.xlsm'!VerifyProductSelectorPosition")
    Write-Output $result
    if ($result -notmatch 'Aligned=(True|Vrai)' -or $result -notmatch 'Choices=4') { throw 'Selector position verification failed' }
} finally {
    $wb.VBProject.VBComponents.Remove($helper)
    $wb.Save()
}
