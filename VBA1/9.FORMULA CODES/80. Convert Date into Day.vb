Sub date2day()
Dim tempCell As Range
Selection.Value = Selection.Value
For Each tempCell In Selection
If IsDate(tempCell) = True Then
With tempCell
.Value = Day(tempCell)
.NumberFormat = "0"
End With
End If
Next tempCell
End Sub