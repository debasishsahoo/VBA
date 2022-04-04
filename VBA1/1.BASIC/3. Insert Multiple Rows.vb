Sub InsertMultipleRows()
Dim i As Integer
Dim j As Integer
ActiveCell.EntireRow.Select
On Error GoTo Last
i = InputBox("Enter number of columns to insert", "Insert Columns")
For j = 1 To i
Selection.Insert Shift:=xlToDown, CopyOrigin:=xlFormatFromRightorAbove
Next j
Last: Exit Sub
End Sub