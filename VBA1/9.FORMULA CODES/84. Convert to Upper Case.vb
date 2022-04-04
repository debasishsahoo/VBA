Sub convertUpperCase()
Dim Rng As Range
For Each Rng In Selection
If Application.WorksheetFunction.IsText(Rng) Then
Rng.Value = UCase(Rng)
End If
Next
End Sub