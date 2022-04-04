Sub SaveWorkshetAsPDF()
Dimws As Worksheet
For Each ws In Worksheets
ws.ExportAsFixedFormat _
xlTypePDF, _
"ENTER-FOLDER-NAME-HERE" & _
ws.Name & ".pdf"
Next ws
End Sub