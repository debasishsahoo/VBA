Sub ForEachDisplayWorkbookNames()
Dim workBookNames As String

For Each wrkbook In Workbooks
workBookNames = workBookNames & vbNewLine & wrkbook.Name
Next wrkbook

MsgBox "The Workbook names are : " & vbNewLine & workBookNames
End Sub