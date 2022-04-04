Sub convertTextCase()
Dim Rng As Range
For Each Rng In Selection
If WorksheetFunction.IsText(Rng) Then
Rng.Value = UCase(Left(Rng, 1)) & LCase(Right(Rng, Len(Rng) - 1))
End If
Next Rng
End Sub