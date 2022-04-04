Sub deleteBlankWorksheets()
Dim Ws As Worksheet
On Error Resume Next
Application.ScreenUpdating= False
Application.DisplayAlerts= False
For Each Ws In Application.Worksheets
If Application.WorksheetFunction.CountA(Ws.UsedRange) = 0 Then
Ws.Delete
End If
Next
Application.ScreenUpdating= True
Application.DisplayAlerts= True
End Sub