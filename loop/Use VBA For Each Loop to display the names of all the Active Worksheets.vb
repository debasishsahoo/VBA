Sub ForEachDisplaySheetNames()
Dim sheetNames As String
For Each sht In ActiveWorkbook.Sheets
sheetNames = sheetNames & vbNewLine & sht.Name
Next sht

MsgBox "The Sheet names are : " & vbNewLine & sheetNames
End Sub