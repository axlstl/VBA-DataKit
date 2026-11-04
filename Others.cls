VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Others"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub SetMacroMode()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

 

Sub DisableMacroMode()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

 

Function Get_WS_Codename(ByVal sName As String) As Worksheet
    Dim lCount As Long

    For lCount = 1 To ThisWorkbook.Worksheets.Count
        If sName = ThisWorkbook.Worksheets(lCount).CodeName Then
            Set Get_WS_Codename = ThisWorkbook.Worksheets(lCount)
            Exit For
        End If
    Next lCount
End Function
