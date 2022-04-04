Sub getCubeRoot()
Dim rng As Range
Dimi As Integer
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = rng ^ (1 / 3)
Else
End If
Nextrng
End Sub