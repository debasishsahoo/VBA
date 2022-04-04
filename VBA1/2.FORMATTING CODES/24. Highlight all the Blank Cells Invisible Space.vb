Sub blankWithSpace()
Dim rng As Range
For Each rng In ActiveSheet.UsedRange
If rng.Value = " " Then
rng.Style = "Note"
End If
Next rng
End Sub