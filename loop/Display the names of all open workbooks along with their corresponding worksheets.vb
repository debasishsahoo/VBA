Sub ForEachLoopNesting()
Dim result As String
For Each wrkbook In Workbooks
For Each sht In wrkbook.Sheets
result = result & vbNewLine & " Workbook : " & wrkbook.Name & " Worksheet : " & sht.Name
Next sht
Next wrkbook

MsgBox result
End Sub