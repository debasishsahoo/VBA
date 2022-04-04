Sub InsertMultipleSheets()
Dim i As Integer
i = _
InputBox("Enter number of sheets to insert.", _
"Enter Multiple Sheets")
Sheets.Add After:=ActiveSheet, Count:=i
End Sub