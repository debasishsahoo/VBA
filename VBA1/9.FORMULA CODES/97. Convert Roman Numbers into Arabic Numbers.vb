Sub convertToNumbers()
Dim rng As Range
Selection.Value = Selection.Value
For Each rng In Selection
If Not WorksheetFunction.IsNonText(rng) Then
rng.Value = WorksheetFunction.Arabic(rng)
End If
Next rng
End Sub