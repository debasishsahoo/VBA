Sub DeleteWorksheets()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If ws.name <> ThisWorkbook.ActiveSheet.name Then
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
End If
Next ws
End Sub