Sub replaceBlankWithZero()
Dim rng As Range
Selection.Value = Selection.Value
For Each rng In Selection
If rng = "" Or rng = " " Then
rng.Value = "0"
Else
End If
Next rng
End Sub