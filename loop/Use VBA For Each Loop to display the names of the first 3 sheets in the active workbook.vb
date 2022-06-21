Sub ForEachDisplayFirstThreeSheetNames()
Dim sheetNames As String
Dim sheetCounter As Integer

For Each sht In ActiveWorkbook.Sheets
sheetNames = sheetNames & vbNewLine & sht.Name
sheetCounter = sheetCounter + 1

If sheetCounter >= 3 Then
Exit For
End If
Next sht

MsgBox "The Sheet names are : " & vbNewLine & sheetNames
End Sub