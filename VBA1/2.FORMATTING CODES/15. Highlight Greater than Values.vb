Sub HighlightGreaterThanValues()
Dim i As Integer
i = InputBox("Enter Greater Than Value", "Enter Value")
Selection.FormatConditions.Delete
Selection.FormatConditions.Add Type:=xlCellValue, _
Operator:=xlGreater, Formula1:=i
Selection.FormatConditions(Selection.FormatConditions.Count).S
tFirstPriority
With Selection.FormatConditions(1)
.Font.Color = RGB(0, 0, 0)
.Interior.Color = RGB(31, 218, 154)
End With
End Sub