Sub SortWorksheets()
Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult
iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
& "Clicking No will sort in Descending Order", _
vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
For i = 1 To Sheets.Count
For j = 1 To Sheets.Count - 1
If iAnswer = vbYes Then
If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
Sheets(j).Move After:=Sheets(j + 1)
End If
ElseIf iAnswer = vbNo Then
If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then Sheets(j).Move After:=Sheets(j + 1)
End If
End If
Next j
Next i
End Sub