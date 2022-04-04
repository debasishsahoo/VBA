Sub removeTime()
Dim Rng As Range
For Each Rng In Selection
If IsDate(Rng) = True Then
Rng.Value = VBA.Int(Rng.Value)
End If
Next
Selection.NumberFormat = "dd-mmm-yy"
End Sub