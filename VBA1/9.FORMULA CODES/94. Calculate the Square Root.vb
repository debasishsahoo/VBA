Sub getSquareRoot()
Dim rng As Range
Dim i As Integer
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = Sqr(rng)
Else
End If
Next rng
End Sub