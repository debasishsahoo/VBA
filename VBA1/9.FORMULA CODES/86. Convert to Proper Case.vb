Sub convertProperCase()
Dim Rng As Range
For Each Rng In Selection
If WorksheetFunction.IsText(Rng) Then
Rng.Value = WorksheetFunction.Proper(Rng.Value)
End If
Next
End Sub