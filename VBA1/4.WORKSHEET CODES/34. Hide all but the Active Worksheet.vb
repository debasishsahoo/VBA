Sub HideWorksheet()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
ws.Visible = xlSheetHidden
End If
Next ws
End Sub