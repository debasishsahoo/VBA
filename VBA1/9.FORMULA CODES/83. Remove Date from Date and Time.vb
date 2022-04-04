Sub removeDate()
Dim Rng As Range
For Each Rng In Selection
If IsDate(Rng) = True Then
Rng.Value = Rng.Value - VBA.Fix(Rng.Value)
End If
NextSelection.NumberFormat = "hh:mm:ss am/pm"
End Sub