Sub HighlightDuplicateValues()
Dim myRange As Range
Dim myCell As Range
Set myRange = Selection
For Each myCell In myRange
If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
myCell.Interior.ColorIndex = 36
End If
Next myCell
End Sub