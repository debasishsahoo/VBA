Sub highlightNegativeNumbers()
Dim Rng As Range
For Each Rng In Selection
If WorksheetFunction.IsNumber(Rng) Then
If Rng.Value < 0 Then
Rng.Font.Color= -16776961
End If
End If
Next
End Sub