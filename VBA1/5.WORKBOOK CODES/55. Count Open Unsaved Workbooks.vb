Sub VisibleWorkbooks()
Dim book As Workbook
Dim i As Integer
For Each book In Workbooks
If book.Saved = False Then
i = i + 1
End If
Next book
MsgBox i
End Sub