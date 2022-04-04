Sub TimeStamp()
Dim i As Integer
For i = 1 To 24
ActiveCell.FormulaR1C1 = i & ":00"
ActiveCell.NumberFormat = "[$-409]h:mm AM/PM;@"
ActiveCell.Offset(RowOffset:=1, ColumnOffset:=0).Select
Next i
End Sub