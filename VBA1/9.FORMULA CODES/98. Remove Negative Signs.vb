Sub removeNegativeSign()
Dim rng As Range
Selection.Value = Selection.Value
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = Abs(rng)
End If
Next rng