Sub highlightAlternateRows()
Dim rng As Range
For Each rng In Selection.Rows
If rng.Row Mod 2 = 1 Then
rng.Style = "20% -Accent1"
rng.Value = rng ^ (1 / 3)
Else
End If
Next rng
End Sub